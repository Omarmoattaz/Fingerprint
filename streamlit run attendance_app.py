import streamlit as st
import pandas as pd
from io import BytesIO

def process_attendance(df):
    # تعديل أسماء الأعمدة لتتناسب مع ما في ملفك
    df = df.rename(columns=lambda x: x.strip())
    
    # تأكد من نوع التاريخ والوقت
    df['Date'] = pd.to_datetime(df['Date'], dayfirst=True).dt.date
    df['Time'] = pd.to_datetime(df['Time']).dt.time

    # دمج التاريخ والوقت في datetime
    df['DateTime'] = pd.to_datetime(df['Date'].astype(str) + ' ' + df['Time'].astype(str))

    # ترتيب البيانات حسب الاسم والتاريخ والوقت
    df = df.sort_values(by=['Name', 'Date', 'DateTime'])

    results = []
    summary = {}

    employees = df['Name'].unique()

    for emp in employees:
        emp_data = df[df['Name'] == emp]

        days = emp_data['Date'].unique()
        total_work_seconds = 0
        attendance_days = 0
        in_out_count = 0
        missing_days = []

        # لمعرفة كل الأيام الموجودة في الملف (يمكن تعديلها لتكون فترة محددة)
        all_days = pd.date_range(start=df['Date'].min(), end=df['Date'].max()).date

        for day in all_days:
            day_records = emp_data[emp_data['Date'] == day]

            if day_records.empty:
                missing_days.append(str(day))
                continue

            in_times = day_records[day_records['Status'].str.contains('C/In', case=False)]['DateTime']
            out_times = day_records[day_records['Status'].str.contains('C/Out', case=False)]['DateTime']

            in_out_count += len(day_records)

            if in_times.empty or out_times.empty:
                # يوم فيه دخول أو خروج مفقود
                results.append({
                    'Name': emp,
                    'Date': day,
                    'First In': in_times.min() if not in_times.empty else None,
                    'Last Out': out_times.max() if not out_times.empty else None,
                    'Work Hours': None
                })
                continue

            first_in = in_times.min()
            last_out = out_times.max()

            work_seconds = (last_out - first_in).total_seconds()
            total_work_seconds += work_seconds
            attendance_days += 1

            hours = int(work_seconds // 3600)
            minutes = int((work_seconds % 3600) // 60)
            work_hours_str = f"{hours}h {minutes}m"

            results.append({
                'Name': emp,
                'Date': day,
                'First In': first_in.time(),
                'Last Out': last_out.time(),
                'Work Hours': work_hours_str
            })

        total_hours = total_work_seconds / 3600
        summary[emp] = {
            'Total Work Hours': round(total_hours, 2),
            'Attendance Days': attendance_days,
            'In/Out Records Count': in_out_count,
            'Missing Days': missing_days
        }

    return pd.DataFrame(results), summary

def to_excel(df, summary):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Daily Attendance')

        # ملخص الموظفين
        summary_df = pd.DataFrame([
            {
                'Name': k,
                'Total Work Hours': v['Total Work Hours'],
                'Attendance Days': v['Attendance Days'],
                'In/Out Records Count': v['In/Out Records Count'],
                'Missing Days': ', '.join(v['Missing Days']) if v['Missing Days'] else '-'
            }
            for k, v in summary.items()
        ])
        summary_df.to_excel(writer, index=False, sheet_name='Summary')

        writer.save()
        processed_data = output.getvalue()
    return processed_data

st.title("تطبيق حساب حضور وانصراف الموظفين")

uploaded_file = st.file_uploader("اختر ملف بصمة الموظفين (Excel)", type=['xlsx', 'xls'])

if uploaded_file:
    df = pd.read_excel(uploaded_file)
    st.write("البيانات الأصلية من الملف:")
    st.dataframe(df.head())

    if st.button("معالجة البيانات وحساب الحضور"):
        with st.spinner("جارٍ معالجة البيانات..."):
            daily_df, summary = process_attendance(df)

        st.success("تمت المعالجة بنجاح!")

        st.subheader("جدول الحضور اليومي")
        st.dataframe(daily_df)

        st.subheader("ملخص الحضور لكل موظف")
        summary_display = pd.DataFrame([
            {
                'Name': k,
                'Total Work Hours': v['Total Work Hours'],
                'Attendance Days': v['Attendance Days'],
                'In/Out Records Count': v['In/Out Records Count'],
                'Missing Days': ', '.join(v['Missing Days']) if v['Missing Days'] else '-'
            }
            for k, v in summary.items()
        ])
        st.dataframe(summary_display)

        excel_data = to_excel(daily_df, summary)
        st.download_button(
            label="تحميل النتائج Excel",
            data=excel_data,
            file_name="attendance_summary.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
