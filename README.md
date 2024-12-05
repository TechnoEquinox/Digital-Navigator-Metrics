<h1>Digital Navigator Metrics</h1>

<h3>Wallingford Public Library: Tech Connect Program</h3>

![image1](https://github.com/user-attachments/assets/da4d5837-9b13-4a9a-b34e-33bd94c05830)

<h2>About</h2>
This script analyzes the Appointment Summary (Responses) workbook to provide metrics for
Tech Connect employees. This script calculates the total appointments seen
by the entire department for the month, quarter, and year. Additionally, this script
calculates an individual employee's total appointments for the month, quarter, and year. 
Appointments are tracked based on the location where they occur, 
and employee appointment pace is measured per month, quarter, and year.


<h2>Limitations</h2>

- Employee names currently need to be manually added in Code.gs
- The script assumes the data in the Google Sheet has already been properly configured
- The script also needs proper configuration to protect PII (Personally Identifiable Information), ensuring only Digital Navigators can access this information
- This script is used along side an existing web page configured with Google Sites. Once the script is deployed via Google Scripts, it should be embedded in the website


<h2>Feature Request</h2>

- Calculate the top five most popular "main goals" that clients bring for all employees
- Track the clients that are seen the most. Analyze where their appointments are taken and what is done in these appointments
- Graph the appointments seen by location in a stacked bar chart
- Calculate the appointments seen per location for each individual employee
