import pandas as pd
import streamlit as st
from datetime import timedelta, date

class BaseForm:
    def __init__(self, name, hrt, interval, start_date):
        self.name = name
        self.hrt = hrt
        self.interval = interval
        self.start_date = start_date

    def generate_calendar(self):
        raise NotImplementedError("Subclasses should implement this method!")

    def create_common_inputs(self):
        self.name = st.text_input("Your Name")
        self.ester = st.text_input("Your HRT")
        self.interval = st.number_input("Dose Interval (days)", min_value=1, format="%d")
        self.start_date = st.date_input("Start Date")

class DosageForm(BaseForm):
    def __init__(self):
        super().__init__(None, None, None, None)
        self.dose = None
        self.total_volume = None
        self.concentration = None

    def create_inputs(self):
        self.create_common_inputs()
        self.dose = st.number_input("Your Dose (mg)", min_value=0.0, format="%.2f")
        self.total_volume = st.number_input("Total Volume of Vial (mL)", min_value=0.0, format="%.2f")
        self.concentration = st.number_input("Concentration of Vial (mg/mL)", min_value=0.0, format="%.2f")

    def generate_calendar(self, date_format, starting_side):
        # Dictionary to store data
        data_dict = {
            "Injection Count": [],
            "Day": [],
            "Date": [],
            "Remaining mL": [],
            "Days": [],
            "Months": [],
            "Side": []
        }

        # Variable initialization
        injection_count = 1
        current_side = starting_side
        days_elapsed = 0
        ml_remaining = self.total_volume

        # Loop to generate the calendar data and store it in the dictionary
        while ml_remaining >= (self.dose / self.concentration):
            injection_date = self.start_date + timedelta(days=days_elapsed)
            dose_ml = self.dose / self.concentration

            # Append data to the dictionary
            data_dict["Injection Count"].append(injection_count)
            data_dict["Day"].append(injection_date.strftime("%A"))
            data_dict["Date"].append(injection_date.strftime(date_format))
            data_dict["Remaining mL"].append(round(ml_remaining, 2))
            data_dict["Days"].append(days_elapsed)
            data_dict["Months"].append(round(days_elapsed / 30, 2))
            data_dict["Side"].append(current_side)

            # Update variables for the next iteration
            ml_remaining -= dose_ml
            days_elapsed += self.interval
            current_side = "Right" if current_side == "Left" else "Left"
            injection_count += 1

        # Convert the dictionary to a DataFrame and display it
        metadata_df = generate_metadata(self.name, date_format, self.dose, self.interval, self.start_date, injection_count, self.total_volume, self.concentration)
        #test = generate_metadata_byobj(self, date_format, injection_count)
        df = pd.DataFrame(data_dict)
        #st.dataframe(test)
        st.dataframe(metadata_df)
        st.dataframe(df)


class VolumeForm(BaseForm):
    def __init__(self):
        super().__init__(None, None, None, None)
        self.injection_amount = None
        self.total_volume = None

    def create_inputs(self):
        self.create_common_inputs()
        self.injection_amount = st.number_input("Injection Amount (mL)", min_value=0.0, format="%.2f")
        self.total_volume = st.number_input("Total Volume of Vial (mL)", min_value=0.0, format="%.2f")

    def generate_calendar(self, date_format, starting_side):
        # Dictionary to store data
        data_dict = {
            "Injection Count": [],
            "Day": [],
            "Date": [],
            "Remaining mL": [],
            "Days": [],
            "Months": [],
            "Side": []
        }

        # Variable initialization
        injection_count = 1
        current_side = starting_side
        days_elapsed = 0
        ml_remaining = self.total_volume

        # Loop to generate the calendar data and store it in the dictionary
        while ml_remaining >= self.injection_amount:
            injection_date = self.start_date + timedelta(days=days_elapsed)

            # Append data to the dictionary
            data_dict["Injection Count"].append(injection_count)
            data_dict["Day"].append(injection_date.strftime("%A"))
            data_dict["Date"].append(injection_date.strftime(date_format))
            data_dict["Remaining mL"].append(round(ml_remaining, 2))
            data_dict["Days"].append(days_elapsed)
            data_dict["Months"].append(round(days_elapsed / 30, 2))
            data_dict["Side"].append(current_side)

            # Update variables for the next iteration
            ml_remaining -= self.injection_amount
            days_elapsed += self.interval
            current_side = "Right" if current_side == "Left" else "Left"
            injection_count += 1

        # Convert the dictionary to a DataFrame and display it

        df = pd.DataFrame(data_dict)
        st.dataframe(df)

def get_date_format(date_format):
    if date_format == "DD/MM/YYYY":
        format = "%d/%m/%Y"
    elif date_format == "MM/DD/YYYY":
        format = "%m/%d/%Y"
    else:
        format = "%Y/%m/%d"
    return format

def generate_metadata(name, date_format, dose, interval, start_date, lasts, total_ml, concentration):
    metadata = {
        "Name": name,
        "Today:": f"{date.today().strftime(date_format)}",
        "Dose (mg)": dose,
        "Dose (mL)": dose / concentration,
        "Interval": interval,
        "Start Date": start_date.strftime(date_format),
        "Lasts": f"{lasts} doses",
        "Total ML": total_ml,
        "Concentration (mg/mL)": concentration
    }
    metadata_df = pd.DataFrame([metadata])
    return metadata_df

def generate_metadata_byobj(dosetype, obj, date_format, injections_sum):
    if dosetype == "mg":
        dosemg = obj.dose
        doseml = obj.dose / obj.concentration
    else:
        dosemg = obj.dose / obj.concentration
    metadata = {
        "Name": obj.name,
        "Today:": f"{date.today().strftime(date_format)}",
        "Dose (mg)": f"{obj.dose} mg",
        "Dose (mL)": f"{obj.dose / obj.concentration} mL",
        "Interval": obj.interval,
        "Start Date": obj.start_date.strftime(date_format),
        "Lasts": f"{injections_sum} doses",
        "Total mL": f"{obj.total_volume} mL",
        "Concentration (mg/mL)": f"{obj.concentration} mg/mL"
    }
    metadata_df = pd.DataFrame([metadata])
    return metadata_df

def main():
    st.sidebar.title("HRT Calendar Generator")
    tab = st.sidebar.radio("Choose Calculation Method", ["By Dosage", "By Injection Amount"])
    date_format = st.sidebar.radio("Choose Date Format", ["DD/MM/YYYY", "MM/DD/YYYY", "YYYY/MM/DD"])
    starting_side = st.sidebar.radio("Choose Starting Side", ["Left", "Right"])

    if tab == "By Dosage":
        st.title("By Dosage")
        dosage_form = DosageForm()
        dosage_form.create_inputs()
        if st.button("Generate Dosage (mg) Calendar"):
            date_format = get_date_format(date_format)
            #st.dataframe(generate_metadata())
            dosage_form.generate_calendar(date_format=date_format, starting_side=starting_side)
    elif tab == "By Injection Amount":
        st.title("By Injection Amount")
        volume_form = VolumeForm()
        volume_form.create_inputs()
        if st.button("Generate Dosage (mL) Calendar"):
            date_format = get_date_format(date_format)
            volume_form.generate_calendar(date_format=date_format, starting_side=starting_side)

# Run the app
main()
