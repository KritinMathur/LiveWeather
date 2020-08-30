## About The Project
\
\
\
![Image of Main](https://github.com/KritinMathur/LiveWeather/blob/master/bin/main.JPG)
\
\
\
Live weather displaying application using APIs and excel to display and control the data being presented to the user. A python script updates the excel sheet temprature and humidity in real time. The other fields apart from 'city name' are required fields. the program uses open weather api to request real time temprature and humidity from web.

## Getting Started

There are 2 methods to get the program running.
\
\
The user can either directly download the project.xlsx which contains 2 sheets. The Main Sheet contains some sample tuples and the List of all city sheet contains the list of all cities and their respective ID.
\
\
The user can also download the city.list.json and setup.py to automatically construct project.xlsx. This will be expanded upon later.

### Prerequisites
```sh
pip install xlwings
```
```sh
pip install openpyxl
```
```sh
pip install requests
```
### Installation
#### Method 1 :-
1. Clone the repository
```sh
git clone https://github.com/KritinMathur/LiveWeather.git
```

#### Method 2 :-
1. Clone the repository
```sh
git clone https://github.com/KritinMathur/LiveWeather.git
```
2. Download the city data json file .
https://drive.google.com/file/d/1UQDzdVogI57NwUL16SdjYKdteRkNLIJy/view?usp=sharing
3. Place city.list.json in bin folder.
4. run setup.py.
\
(This will create a fresh project.xlsx)

### Usage
1. Open the project.xlsx and go to sheet ('list of all cities').
2. select the city and its corresponding city ID for which live weather is to be reported.
3. Copy and paste the city and city ID under their respective columns in sheet('Main Sheet').
4. Enter the unit(F for fahrenheit and C for Celsius) and updation factor (0 to stop updation for that row and 1 to resume/start updating)
5. Run final.py present in the root folder.
6. Enter the updation time interval 
\
(select value > 5 because the API key limits the request rate under 60requests/min and 1000requests/month)
\
7. The temperature and humidity columns will automatically update till the script is force closed with a keyboard inturupt or if the process is terminated.
