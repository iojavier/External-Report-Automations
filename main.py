import streamlit as st
import pandas as pd
import datetime
from io import BytesIO
from pandas import ExcelWriter
import uuid

st.set_page_config(layout="wide", page_title="Daily Remark Summary", page_icon="📊", initial_sidebar_state="expanded")
st.title('Daily Remark Summary')

@st.cache_data
def load_data(uploaded_file):
    df = pd.read_excel(uploaded_file)
    df.columns = df.columns.str.strip().str.upper()
    df['DATE'] = pd.to_datetime(df['DATE'], errors='coerce')
    df = df[df['DATE'].dt.weekday != 6]
    return df

def to_excel(summary_groups):
    output = BytesIO()
    with ExcelWriter(output, engine='xlsxwriter', date_format='yyyy-mm-dd') as writer:
        workbook = writer.book
        formats = {
            'title': workbook.add_format({'bold': True, 'font_size': 14, 'align': 'center', 'valign': 'vcenter', 'bg_color': '#FFFF00'}),
            'center': workbook.add_format({'align': 'center', 'valign': 'vcenter', 'border': 1}),
            'header': workbook.add_format({'align': 'center', 'valign': 'vcenter', 'bg_color': 'red', 'font_color': 'white', 'bold': True}),
            'comma': workbook.add_format({'align': 'center', 'valign': 'vcenter', 'border': 1, 'num_format': '#,##0'}),
            'percent': workbook.add_format({'align': 'center', 'valign': 'vcenter', 'border': 1, 'num_format': '0.00%'}),
            'date': workbook.add_format({'align': 'center', 'valign': 'vcenter', 'border': 1, 'num_format': 'yyyy-mm-dd'}),
            'time': workbook.add_format({'align': 'center', 'valign': 'vcenter', 'border': 1, 'num_format': 'hh:mm:ss'})
        }

        for sheet_name, df_dict in summary_groups.items():
            worksheet = workbook.add_worksheet(sheet_name)
            current_row = 0

            for title, df in df_dict.items():
                if df.empty:
                    continue
                df_for_excel = df.copy()
                for col in ['PENETRATION RATE (%)', 'CONNECTED RATE (%)', 'PTP RATE', 'CALL DROP RATIO #']:
                    df_for_excel[col] = df_for_excel[col].str.rstrip('%').astype(float)

                # Write title
                worksheet.merge_range(current_row, 0, current_row, len(df.columns) - 1, title, formats['title'])
                current_row += 1

                # Write headers
                for col_num, col_name in enumerate(df_for_excel.columns):
                    worksheet.write(current_row, col_num, col_name, formats['header'])
                    max_len = max(df_for_excel[col_name].astype(str).str.len().max(), len(col_name)) + 2
                    worksheet.set_column(col_num, col_num, max_len)

                current_row += 1

                # Write data
                for row_num in range(len(df_for_excel)):
                    for col_num, col_name in enumerate(df_for_excel.columns):
                        value = df_for_excel.iloc[row_num, col_num]
                        if col_name == 'DATE':
                            worksheet.write_datetime(current_row + row_num, col_num, value, formats['date'])
                        elif col_name in ['TOTAL PTP AMOUNT', 'TOTAL BALANCE']:
                            worksheet.write(current_row + row_num, col_num, value, formats['comma'])
                        elif col_name in ['PENETRATION RATE (%)', 'CONNECTED RATE (%)', 'PTP RATE', 'CALL DROP RATIO #']:
                            worksheet.write(current_row + row_num, col_num, value / 100, formats['percent'])
                        elif col_name in ['TOTAL TALK TIME', 'TALK TIME AVE']:
                            worksheet.write(current_row + row_num, col_num, value, formats['time'])
                        else:
                            worksheet.write(current_row + row_num, col_num, value, formats['center'])

                current_row += len(df_for_excel) + 2  # Add 2 empty rows as spacer

    return output.getvalue()

uploaded_files = st.sidebar.file_uploader("Upload Daily Remark Files", type="xlsx", accept_multiple_files=True)

def process_file(df):
    df = df[df['REMARK BY'] != 'SPMADRID']
    df = df[~df['DEBTOR'].str.contains("DEFAULT_LEAD_", case=False, na=False)]
    df = df[~df['STATUS'].str.contains('ABORT', na=False)]
    df = df[~df['REMARK'].str.contains(r'1_\d{11} - PTP NEW', case=False, na=False, regex=True)]
    excluded_remarks = ["Broken Promise", "New files imported", "Updates when case reassign to another collector", 
                        "NDF IN ICS", "FOR PULL OUT (END OF HANDLING PERIOD)", "END OF HANDLING PERIOD", "New Assignment -", "broadcast", "File Unhold"]
    df = df[~df['REMARK'].str.contains('|'.join(excluded_remarks), case=False, na=False)]
    df = df[~df['CALL STATUS'].str.contains('OTHERS', case=False, na=False)]
    df['CARD NO.'] = df['CARD NO.'].astype(str)
    df['CYCLE'] = df['CARD NO.'].str[:2].fillna('Unknown')
    return df

def format_seconds_to_hms(seconds):
    seconds = int(seconds)
    hours, minutes = seconds // 3600, (seconds % 3600) // 60
    return f"{hours:02d}:{minutes:02d}:{seconds % 60:02d}"

def calculate_summary(df, remark_types, manual_correction=False):
    summary_columns = ['DATE', 'CLIENT', 'COLLECTORS', 'ACCOUNTS', 'TOTAL DIALED', 'PENETRATION RATE (%)', 
                      'CONNECTED #', 'CONNECTED RATE (%)', 'CONNECTED ACC', 'TOTAL TALK TIME', 'TALK TIME AVE', 
                      'CONNECTED AVE', 'PTP ACC', 'PTP RATE', 'TOTAL PTP AMOUNT', 'TOTAL BALANCE', 'CALL DROP #', 
                      'SYSTEM DROP', 'CALL DROP RATIO #']
    summary_table = pd.DataFrame(columns=summary_columns)
    df_filtered = df[df['REMARK TYPE'].isin(remark_types)].copy()
    df_filtered['DATE'] = df_filtered['DATE'].dt.date

    for (date, client), group in df_filtered.groupby(['DATE', 'CLIENT']):
        collectors = group[group['CALL DURATION'].notna()]['REMARK BY'].nunique()
        if collectors == 0:
            continue
        accounts = group['ACCOUNT NO.'].nunique()
        total_dialed = group['ACCOUNT NO.'].count()
        connected = group[group['CALL STATUS'] == 'CONNECTED']['ACCOUNT NO.'].nunique()
        penetration_rate = f"{(total_dialed / accounts * 100):.2f}%" if accounts else "0.00%"
        connected_acc = group[group['CALL STATUS'] == 'CONNECTED']['ACCOUNT NO.'].count()
        connected_rate = f"{(connected_acc / total_dialed * 100):.2f}%" if total_dialed else "0.00%"
        ptp_acc = group[(group['STATUS'].str.contains('PTP', na=False)) & (group['PTP AMOUNT'] != 0)]['ACCOUNT NO.'].nunique()
        ptp_rate = f"{(ptp_acc / connected * 100):.2f}%" if connected else "0.00%"
        total_ptp_amount = group[(group['STATUS'].str.contains('PTP', na=False)) & (group['PTP AMOUNT'] != 0)]['PTP AMOUNT'].sum()
        total_balance = group[(group['PTP AMOUNT'] != 0)]['BALANCE'].sum()
        system_drop = group[(group['STATUS'].str.contains('DROPPED', na=False)) & (group['REMARK BY'] == 'SYSTEM')]['ACCOUNT NO.'].count()
        call_drop_count = group[(group['STATUS'].str.contains('NEGATIVE CALLOUTS - DROP CALL|NEGATIVE_CALLOUTS - DROPPED_CALL', na=False)) & 
                              (~group['REMARK BY'].str.upper().isin(['SYSTEM']))]['ACCOUNT NO.'].count()
        call_drop_ratio = f"{(call_drop_count / connected_acc * 100):.2f}%" if manual_correction and connected_acc else \
                         f"{(system_drop / connected_acc * 100):.2f}%" if connected_acc else "0.00%"
        total_talk_seconds = group['TALK TIME DURATION'].sum()
        total_talk_time = format_seconds_to_hms(total_talk_seconds)
        talk_time_ave = format_seconds_to_hms(total_talk_seconds / collectors) if collectors else "00:00:00"
        connected_ave = round(connected_acc / collectors, 2) if collectors else 0

        summary_table = pd.concat([summary_table, pd.DataFrame([{
            'DATE': date, 'CLIENT': client, 'COLLECTORS': collectors, 'ACCOUNTS': accounts, 'TOTAL DIALED': total_dialed,
            'PENETRATION RATE (%)': penetration_rate, 'CONNECTED #': connected, 'CONNECTED RATE (%)': connected_rate,
            'CONNECTED ACC': connected_acc, 'TOTAL TALK TIME': total_talk_time, 'TALK TIME AVE': talk_time_ave,
            'CONNECTED AVE': connected_ave, 'PTP ACC': ptp_acc, 'PTP RATE': ptp_rate, 'TOTAL PTP AMOUNT': total_ptp_amount,
            'TOTAL BALANCE': total_balance, 'CALL DROP #': call_drop_count, 'SYSTEM DROP': system_drop, 'CALL DROP RATIO #': call_drop_ratio
        }])], ignore_index=True)
    
    return summary_table.sort_values(by=['DATE'])

def get_cycle_summary(df, remark_types, manual_correction=False):
    result = {}
    for cycle in df['CYCLE'].unique():
        if cycle.lower() in ['unknown', 'na']:
            continue
        cycle_df = df[df['CYCLE'] == cycle]
        result[f"Cycle {cycle}"] = calculate_summary(cycle_df, remark_types, manual_correction)
    return result

def get_balance_summary(df, remark_types, manual_correction=False):
    balance_ranges = [
        (0, 9999.99, "0-9999.99"),
        (10000.00, 49999.99, "10000.00-49999.99"),
        (50000.00, 99999.99, "50000.00-99999.99"),
        (100000.00, float('inf'), "100000.00 and up")
    ]
    result = {}
    for cycle in df['CYCLE'].unique():
        if cycle.lower() in ['unknown', 'na']:
            continue
        cycle_df = df[df['CYCLE'] == cycle]
        for min_bal, max_bal, range_name in balance_ranges:
            balance_df = cycle_df[(cycle_df['BALANCE'] >= min_bal) & (cycle_df['BALANCE'] <= max_bal)]
            if not balance_df.empty:
                summary = calculate_summary(balance_df, remark_types, manual_correction)
                result[f"Cycle {cycle} Balance {range_name}"] = summary
    return result

if uploaded_files:
    all_combined = []
    all_predictive = []
    all_manual = []
    predictive_cycle_summaries = {}
    manual_cycle_summaries = {}
    predictive_balance_summaries = {}
    manual_balance_summaries = {}

    for idx, file in enumerate(uploaded_files, 1):
        df = load_data(file)
        df = process_file(df)
        combined_summary = calculate_summary(df, ['Predictive', 'Follow Up', 'Outgoing'])
        predictive_summary = calculate_summary(df, ['Predictive', 'Follow Up'])
        manual_summary = calculate_summary(df, ['Outgoing'], manual_correction=True)
        predictive_cycles = get_cycle_summary(df, ['Predictive', 'Follow Up'])
        manual_cycles = get_cycle_summary(df, ['Outgoing'], manual_correction=True)
        predictive_balances = get_balance_summary(df, ['Predictive', 'Follow Up'])
        manual_balances = get_balance_summary(df, ['Outgoing'], manual_correction=True)

        all_combined.append(combined_summary)
        all_predictive.append(predictive_summary)
        all_manual.append(manual_summary)
        
        for cycle, table in predictive_cycles.items():
            if cycle not in predictive_cycle_summaries:
                predictive_cycle_summaries[cycle] = table
            else:
                predictive_cycle_summaries[cycle] = pd.concat([predictive_cycle_summaries[cycle], table], ignore_index=True)
                
        for cycle, table in manual_cycles.items():
            if cycle not in manual_cycle_summaries:
                manual_cycle_summaries[cycle] = table
            else:
                manual_cycle_summaries[cycle] = pd.concat([manual_cycle_summaries[cycle], table], ignore_index=True)
                
        for balance_range, table in predictive_balances.items():
            if balance_range not in predictive_balance_summaries:
                predictive_balance_summaries[balance_range] = table
            else:
                predictive_balance_summaries[balance_range] = pd.concat([predictive_balance_summaries[balance_range], table], ignore_index=True)
                
        for balance_range, table in manual_balances.items():
            if balance_range not in manual_balance_summaries:
                manual_balance_summaries[balance_range] = table
            else:
                manual_balance_summaries[balance_range] = pd.concat([manual_balance_summaries[balance_range], table], ignore_index=True)
        
        st.write(f"Process Done {idx}")

    combined_summary = pd.concat(all_combined, ignore_index=True).sort_values(by=['DATE'])
    predictive_summary = pd.concat(all_predictive, ignore_index=True).sort_values(by=['DATE'])
    manual_summary = pd.concat(all_manual, ignore_index=True).sort_values(by=['DATE'])

    st.write("## Overall Combined Summary Table")
    st.write(combined_summary)
    st.write("## Overall Predictive Summary Table")
    st.write(predictive_summary)
    st.write("## Overall Manual Summary Table")
    st.write(manual_summary)

    st.write("## Per Cycle Predictive Summary Tables")
    for cycle, table in predictive_cycle_summaries.items():
        if "Cycle na" not in cycle.lower():
            with st.container():
                st.subheader(f"Summary for {cycle}")
                st.write(table.sort_values(by=['DATE']))

    st.write("## Per Cycle Manual Summary Tables")
    for cycle, table in manual_cycle_summaries.items():
        if "Cycle na" not in cycle.lower():
            with st.container():
                st.subheader(f"Summary for {cycle}")
                st.write(table.sort_values(by=['DATE']))

    st.write("## Per Balance Predictive Summary Tables")
    for balance_range, table in predictive_balance_summaries.items():
        with st.container():
            st.subheader(f"Summary for {balance_range}")
            st.write(table.sort_values(by=['DATE']))

    st.write("## Per Balance Manual Summary Tables")
    for balance_range, table in manual_balance_summaries.items():
        with st.container():
            st.subheader(f"Summary for {balance_range}")
            st.write(table.sort_values(by=['DATE']))

    summary_groups = {
        'Combined': {'Combined Summary': combined_summary},
        'Predictive': {'Predictive Summary': predictive_summary},
        'Manual': {'Manual Summary': manual_summary},
        'Predictive Cycles': {f"Cycle {k.split('Cycle ')[1]}": v for k, v in predictive_cycle_summaries.items() if "Cycle na" not in k.lower()},
        'Manual Cycles': {f"Cycle {k.split('Cycle ')[1]}": v for k, v in manual_cycle_summaries.items() if "Cycle na" not in k.lower()},
        'Predictive Balances': {k: v for k, v in predictive_balance_summaries.items()},
        'Manual Balances': {k: v for k, v in manual_balance_summaries.items()}
    }

    st.download_button(
        label="Download All Summaries as Excel",
        data=to_excel(summary_groups),
        file_name=f"Daily_Remark_Summary_{datetime.datetime.now().strftime('%Y%m%d')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
