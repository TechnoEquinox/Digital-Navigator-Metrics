<h1>Digital Navigator Metrics</h1>

<h3>Wallingford Public Library: Tech Connect Program</h3>

![image1](https://github.com/user-attachments/assets/da4d5837-9b13-4a9a-b34e-33bd94c05830)

<h2>About</h2>
This is a collection of scripts and HTML documents used by the Digital Navigators at
the Wallingford Public Library. The project is broken into 3 main components:

<h4>Appointment Dashboard</h4>
This dashboard aggregates the responses filled out by Digital Navigators in their Appointment Summary Form
into digestable data that helps both the digital navigators and their managers make informed decisions. This
script tracks the appointments taken each month by location and by employee and breaks down the relevant 
statistics, such as appointments taken, appointment pace, appointments taken in each location, etc.

<h4>Call Log Dashboard</h4>
This dashboard helps digital navigators analyze their pending calls and voicemails. Digital Navigators 
receive many calls every day from community memebers looking for help. At the Wallingford Library, we
are already logging these calls in a shared Google Sheet. This script uses this sheet as the source for
it's data, and filters through the massive log of calls to highlight the clients that still need to be
contacted. The webpage displays an organized table with the client's name, phone number, reason for calling,
and days since last contact. This allows Digital Navigators to spend less time worrying about clients falling
through the cracks and more time helping the community.


<h4>Client Feedback Dashboard</h4>
This dashboard allows Digital Navigators and their managers to easily view client feedback to the program.
We currently use Aquity Scheduling to manage bookings and clients. Aquity is configured to send an email
to the client 24 hours after their scheduled meeting with a Digital Navigator. This email contains a link
to a Google Form where clients can fill out feedback on their appointment. Our script analyzes this data 
and presents it in a dashboard, allowing Digital Navigators to understand what their clients enjoy most about
working with them, and where clients feel improvement is needed. Managers have easy access to check in on 
their employees, so their metrics can be a part of their yearly review.



<h2>Limitations</h2>

- Employee names currently need to be manually added in Code.gs
- The script assumes the data in the Google Sheet has already been properly configured with appropriate form names, IDs, column names, and timestamps.
- This script is used along side an existing web page configured with Google Sites. Once the script is deployed via Google Scripts, it should be embedded in the website.
- When using these scripts, administrators should ensure that PII (Personally Identifiable Information) is protected with the necessary access control mechanisms.


<h2>Feature Request</h2>

- Calculate the top five most popular "main goals" that clients bring for all employees.
- Track the clients that are seen the most. Analyze where their appointments are taken and what is done in these appointments.
- Calculate the appointments seen per location for each individual employee.
