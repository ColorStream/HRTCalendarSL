# INITIAL IMPORTS
import streamlit as st
from datetime import timedelta, date
import pandas as pd

class Form:
    def __init__(self, name=None, hrt=None, interval=None, start_date=None, dose=None, dosetype=["mg", "mL"], total_volume=None, concentration=None):
        self.name = name
        self.ester = hrt
        self.interval = interval
        self.start_date = start_date
        self.dose = dose
        self.dosetype = dosetype
        self.total_volume = total_volume
        self.concentration = concentration

    def create_inputs(self, date_format):
        self.name = st.text_input("Your Name", value=self.name)
        self.ester = st.text_input("Your [Ester](https://en.wikipedia.org/wiki/Steroid_ester)", value=self.ester)
        self.interval = st.number_input("Dose Interval (days)", min_value=1, format="%d", value=self.interval)
        self.start_date = st.date_input("Start Date", value=self.start_date, format=date_format)

        # dosage inputs
        self.dose = st.number_input(f"Your Dose ({self.dosetype})", min_value=0.0, format="%.2f", value=self.dose)
        self.total_volume = st.number_input("Total Volume of Vial (mL)", min_value=0.0, format="%.2f", value=self.total_volume)
        self.concentration = st.number_input("Concentration of Vial (mg/mL)", min_value=0.0, format="%.2f", value=self.concentration)

    def generate_metadata(self, date_format, injections_sum):
        if self.dosetype == "mg":
            dosemg = self.dose
            doseml = self.dose / self.concentration
        else:
            dosemg = self.dose * self.concentration
            doseml = self.dose
        metadata = {
            "Name": self.name,
            "Ester": self.ester,
            "Generation Date": f"{date.today().strftime(date_format)}",
            "Dose (mg)": f"{dosemg} mg",
            "Dose (mL)": f"{doseml} mL",
            "Interval": self.interval,
            "Start Date": self.start_date.strftime(date_format),
            "Lasts": f"{injections_sum} doses",
            "Total mL": f"{self.total_volume} mL",
            "Concentration (mg/mL)": f"{self.concentration} mg/mL"
        }
        metadata_df = pd.DataFrame([metadata])
        return metadata_df
    
    def validate_fields(self):
        required_fields = [self.name, self.ester, self.interval, self.start_date, self.dose, self.total_volume, self.concentration]
        return all(field is not None and field != '' for field in required_fields)

    def generate_calendar(self, date_format, starting_side):
        # dictionary, since you can't append to a dataframe
        data_dict = {
            "Injection Count": [],
            "Day": [],
            "Date": [],
            "Remaining mL": [],
            "Days": [],
            "Months": [],
            "Side": []
        }

        # variable initialization
        injection_count = 1
        current_side = starting_side
        days_elapsed = 0
        ml_remaining = self.total_volume
        if self.dosetype == "mg":
            dose_ml = self.dose / self.concentration
        else:
            dose_ml = self.dose

        # loop to generate the calendar data and store it in the dictionary
        while ml_remaining >= dose_ml:
            injection_date = self.start_date + timedelta(days=days_elapsed)

            # append data to the dictionary
            data_dict["Injection Count"].append(injection_count)
            data_dict["Day"].append(injection_date.strftime("%A"))
            data_dict["Date"].append(injection_date.strftime(date_format))
            data_dict["Remaining mL"].append(round(ml_remaining, 2))
            data_dict["Days"].append(days_elapsed)
            data_dict["Months"].append(round(days_elapsed / 30, 2))
            data_dict["Side"].append(current_side)

            # update variables for the next iteration
            ml_remaining -= dose_ml
            days_elapsed += self.interval
            current_side = "Right" if current_side == "Left" else "Left"
            injection_count += 1

        #convert the dictionary to a DataFrame and display
        metadata = self.generate_metadata(date_format, injection_count)
        df = pd.DataFrame(data_dict)
        return metadata, df


# OPENPYXL IMPORTS
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl import Workbook
from openpyxl.styles import PatternFill
import io

def export_to_excel(metadata_df, calendar_df):
        if metadata_df is None or calendar_df is None:
            st.error("Metadata or Calendar is not initialized.")
            return

        wb = Workbook()
        ws = wb.active

        try:
            output = io.BytesIO() # have to store in memory, unfortunately
            # write metadata
            for r_idx, row in enumerate(dataframe_to_rows(metadata_df, index=False, header=True), 1):
                for c_idx, value in enumerate(row, 1):
                    ws.cell(row=r_idx, column=c_idx, value=value)
            
            md_fill = PatternFill(start_color='cce6ff', end_color='cce6ff', fill_type='solid') # blue :)
            for cell in ws[1]: #color metadata header
                cell.fill = md_fill

            # write calendar data
            calendar_start_row = len(metadata_df) + 3  # leave 2 rows gap between metadata and calendar
            for r_idx, row in enumerate(dataframe_to_rows(calendar_df, index=False, header=True), calendar_start_row):
                for c_idx, value in enumerate(row, 1):
                    ws.cell(row=r_idx, column=c_idx, value=value)

            cal_fill = PatternFill(start_color='ffccff', end_color='ffccff', fill_type='solid') # pink :)
            for cell in ws[calendar_start_row]: #color calendar header
                cell.fill = cal_fill
            
            # This method is taken from here: https://github.com/Loke-60000/Archaeological-Sites-Sorter/blob/master/app.py
            for col in ws.columns:
                max_length = 0
                column = col[0].column_letter  # get the column letter
                for cell in col:
                    try:
                        if len(str(cell.value)) > max_length:
                            max_length = len(str(cell.value))
                    except:
                        pass
                adjusted_width = max_length + 2  # padding
                ws.column_dimensions[column].width = adjusted_width

            # save to downloads
            wb.save(output)
            output.seek(0)  # Move the cursor to the start of the stream
            
            return output

        except PermissionError:
            return "Error: Unable to save the file. Please close any open copies of the file and try again."
        except Exception as e:
            return f"An unexpected error occurred: {e}"

from icalendar import Calendar, Event
from datetime import datetime

def export_to_ical(calendar_df, warn_secondtolast=False):
    if calendar_df is None:
        st.error("Calendar data is not provided.")
        return

    try:
        output = io.BytesIO() 
        cal = Calendar()
        cal.add('prodid', '-//HRT Calendar//mxm.dk//')
        cal.add('version', '2.0')

        # get the index of the second-to-last dose if the warning is enabled
        second_to_last_index = len(calendar_df) - 2 if warn_secondtolast and len(calendar_df) > 1 else None

        # loop dataframe
        for index, row in calendar_df.iterrows():
            event = Event()

            # default event name
            event.add('summary', f"Take dose: {row['Side']} side")

            # ff this is the second-to-last dose, add a special message
            if warn_secondtolast and index == second_to_last_index:
                event['summary'] = event['summary'] + " (SECOND TO LAST DOSE)"
                event.add('description', "Please order a refill if you haven't already!\n")
            else:
                event.add('description', "")

            description = (f"Injection Count: {row['Injection Count']}\n"
                           f"Remaining mL: {row['Remaining mL']}\n"
                           f"Days: {row['Days']}\n"
                           f"Months: {row['Months']}\n")
            event['description'] = event['description'] + description

            # set the event date
            event_date = datetime.strptime(row['Date'], "%Y-%m-%d")  
            event.add('dtstart', event_date.date())
            event.add('dtend', event_date.date())  # all-day event, so start and end are the same

            cal.add_component(event)

        # save to downloads folder
        output = io.BytesIO()
        output.write(cal.to_ical())
        output.seek(0)  # Move the cursor to the start of the stream

        return output
    
    except PermissionError:
        return "Error: Unable to save the file. Please close any open copies of the file and try again."
    except Exception as e:
        return f"An unexpected error occurred: {e}"

def get_date_format(date_format):
    if date_format == "DD/MM/YYYY":
        format = "%d/%m/%Y"
    elif date_format == "MM/DD/YYYY":
        format = "%m/%d/%Y"
    else:
        format = "%Y/%m/%d"
    return format

def generate_form(dosetype, date_format, starting_side, warn_secondtolast):
    st.title(f"Calendar By Dosage in {dosetype}")
    form = Form(dosetype=dosetype)
    form.create_inputs(date_format=date_format)

    # All the stuff I have to initialize....... I feel like this is bad practice...
    invalidForm = False
    dataframe_display = ""
    exceloutput = ""
    icaloutput = ""

    col1_, col2_, col3_ = st.columns([2,1,2])
    with col1_:
        if st.button(f"Generate Dosage ({dosetype}) Calendar"):
            if not form.validate_fields():
                invalidForm = True
            else:
                date_format = get_date_format(date_format)
                metadata, calendar = form.generate_calendar(date_format=date_format, starting_side=starting_side)
                dataframe_display = [metadata, calendar]
    with col2_:
        if st.button("Export to Excel"):
            if not form.validate_fields():
                invalidForm = True
            else:
                date_format = get_date_format(date_format)
                metadata, calendar = form.generate_calendar(date_format=date_format, starting_side=starting_side)
                exceloutput = export_to_excel(metadata_df=metadata, calendar_df=calendar)
                
    with col3_:
        if st.button("Export to iCal"):
            if not form.validate_fields():
                invalidForm = True
            else:
                warn_secondtolast = False if warn_secondtolast == "No" else True
                metadata, calendar = form.generate_calendar(date_format="%Y-%m-%d", starting_side=starting_side)
                icaloutput = export_to_ical(calendar_df=calendar, warn_secondtolast=warn_secondtolast)

    # Moving things below the columns because it's ugly within them
    if invalidForm:
        st.error("Please fill out all fields properly before exporting the calendar.")
    
    if dataframe_display:
        st.dataframe(dataframe_display[0])
        st.dataframe(dataframe_display[1])
    
    if exceloutput:
        if "Error" in exceloutput or "unexpected" in exceloutput:
            st.error(exceloutput)  
        else:
            st.success("File generated.") 
            st.download_button(
            label="Download Excel",
            data=exceloutput,
            file_name=f"{form.name}_hrt_calendar.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    if icaloutput:
        if "Error" in icaloutput or "unexpected" in icaloutput:
            st.error(icaloutput)
        else:
            st.success("File generated.") 
            st.download_button(
            label="Download iCal",
            data=icaloutput,
            file_name=f"{form.name}_hrt_calendar.ics",
            mime="text/calendar"
            )

def main():
    st.set_page_config(page_title="HRT Calendar Generator", page_icon=None, layout="centered", initial_sidebar_state="auto", menu_items=None)

    # markdown for the button columns
    st.markdown("""
            <style>
                div[data-testid="column"] {
                    width: fit-content !important;
                    flex: unset;
                }
                div[data-testid="column"] * {
                    width: fit-content !important;
                }
            </style>
            """, unsafe_allow_html=True)

    st.sidebar.title("HRT Calendar Generator")
    tab = st.sidebar.radio("Choose Calculation Method", ["By dosage in mg", "By dosage in mL"])
    date_format = st.sidebar.radio("Choose Date Format", ["MM/DD/YYYY", "DD/MM/YYYY", "YYYY/MM/DD"]) #apologies for americanizing this. 
    starting_side = st.sidebar.radio("Choose Starting Side", ["Left", "Right"])
    warn_secondtolast = st.sidebar.radio("Warning at Second to Last Dose? (For ICAL)", ["No", "Yes"])
    st.sidebar.write("Newbie to injectable hormones? [Check out this resource.](https://stainedglasswoman.substack.com/p/what-to-expect-when-youre-injecting)")
    st.sidebar.divider()
    st.sidebar.caption("[Github Repo](https://github.com/ColorStream/HRTCalendarSL)")

    if tab == "By dosage in mg":
        generate_form(dosetype="mg", date_format=date_format, starting_side=starting_side, warn_secondtolast=warn_secondtolast)

    elif tab == "By dosage in mL":
        generate_form(dosetype="mL", date_format=date_format, starting_side=starting_side, warn_secondtolast=warn_secondtolast)

# Run the app
main()
