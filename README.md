# HRT Calendar Streamlit App

This is a fork of [S4r4h-O's HRT Calendar](https://github.com/S4r4h-O/HRT-Calendar-Calend-rio-de-TH), which works offline. I wanted to make a version using [Streamlit](https://streamlit.io/) just so I could host it on the web. Feel free to suggest additions or contribute. 

Access it on Streamlit Community Cloud here: https://hrtcalendar.streamlit.app/

### Pictures
<img width="960" alt="calgen1" src="https://github.com/user-attachments/assets/d2fb1e53-0fb9-4bb3-8016-f54123b7ee6c" />
<img width="960" alt="calgen2" src="https://github.com/user-attachments/assets/ab1a1978-4052-47af-bdf5-0cd02910b215" />

What exporting to Excel looks like:
<img width="960" alt="sheet" src="https://github.com/user-attachments/assets/33ab864e-4d2a-4cf7-91fa-a7d829fc287e" />

What an event after importing the .iCal file looks like:
<br><img width="321" alt="icalexample" src="https://github.com/user-attachments/assets/3c1ce427-354d-42b8-9ef4-8bfbf29424e4" />
(Note: At the moment, I have not implemented the calendar as connected reoccurring events. Events added by the iCal import will all be listed as independent events. I plan on integrating recurring iCal events eventually.)

### Running Locally

If you'd like to run locally, here's a quick run through. You only need Python pre-installed.

1. Fork/clone/download this repository.

```
git clone https://github.com/ColorStream/HRTCalendarSL.git
```

2. Navigate to the directory.

```
cd HRTCalendarSL
```

3. Create a virtual environment for this project. This example uses `.venv` as the name.

```
python -m venv .venv
```

4. Load the virtual environment.
    - On Windows Powershell: `.venv/Scripts/activate.ps1`
    - On Linux and Git Bash: `source venv/bin/activate`

5. Run `pip install -r requirements.txt`.

6. Then run `streamlit run app.py`.

This project is licensed under the terms of the GNU General Public License v3. See the LICENSE file for more details.

Copyright (C) 2025

This program is free software; you may redistribute it and/or modify it
under the terms of the GNU General Public License, as published by the
Free Software Foundation; in version 3 of the License, or (at your option)
any later version.

This program is distributed with the expectation that it will be useful,
but WITHOUT ANY WARRANTY; without even the implied warranty of
MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See
the GNU General Public License for more details.

You should have received a copy of the GNU General Public License
General Public License along with this program. If not, see <https://www.gnu.org/licenses/>.
