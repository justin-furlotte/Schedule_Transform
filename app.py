import streamlit as st
import pandas as pd
import io

st.title("Schedule Processor")

uploaded_file = st.file_uploader("Upload your Excel file (.xlsx)", type=["xlsx"])

if uploaded_file:
    dfs = pd.read_excel(uploaded_file, sheet_name=None)

    df = dfs["FORMATTED"]
    df = df.set_index("time")

    common_activities = pd.DataFrame(index=df.index, columns=["schedule"])
    common_activities.loc["8:25-8:40"] = "Homeroom"
    common_activities.loc["10:10-10:25"] = "Recess"
    common_activities.loc["11:55-12:25"] = "WIN"
    common_activities.loc["12:25-1:25"] = "Lunch"

    teacher_map = dfs["TEACHER_MAP"]
    teacher_map = teacher_map.set_index("value_type")
    teacher_map = teacher_map.T
    teacher_map["value"] = (
            teacher_map.index.astype(str)
            + " - "
            + teacher_map["subject"].astype(str)
            + " - "
            + teacher_map["roomnumber"].astype(str)
    )

    schedules = dict()
    homerooms = [x for x in teacher_map["homeroom"].unique() if
                 isinstance(x, str) and any(char.isdigit() for char in x)]

    for homeroom in homerooms:
        new_df = pd.DataFrame(index=df.index, columns=[homeroom])
        mask = df.astype(str).apply(lambda x: x.str.contains(homeroom, na=False))
        new_df[homeroom] = mask.idxmax(axis=1).where(mask.any(axis=1))

        final_schedule = new_df.combine_first(common_activities.rename(columns={"schedule": homeroom}))
        final_schedule = final_schedule.fillna("PREP")
        final_schedule = final_schedule.map(lambda x: teacher_map.loc[x, "value"] if x in teacher_map.index else x)

        final_schedule = final_schedule.rename(columns={homeroom: "Monday"})
        for day in ["Tuesday", "Wednesday", "Thursday", "Friday"]:
            final_schedule[day] = final_schedule["Monday"]

        schedules[homeroom] = final_schedule

    # Save to in-memory buffer
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        for sheet_name, df in schedules.items():
            df.to_excel(writer, sheet_name=sheet_name.replace("/", "-"), index=False)
    output.seek(0)

    st.success("Wow congrats Jill, you did it!")
    st.download_button(
        label="Download Excel File",
        data=output,
        file_name="output_schedule.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
