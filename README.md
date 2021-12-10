# MICROSOFT TEAMS ATTENDANCE AUTOMATOR

 An application 
that provides easy access to meeting organizers to handle the downloaded attendance list.All the professors schedule classes in their respective time slots and download the attendance list at the end of the class. The downloaded attendance list is a mess and to access it professors have to make a list of presentees and absentees which take a lot of time. 

This attendance automator manifests the meeting organizers to easily sort the absentees and presentees according to their scholar number. It also generates attendance of each student in percentage and highlights it if it is less than the threshold value.

---
## Required Environment Variables

VARIABLE | Sample value
---- | ---
MONGO_URI | create by using MongoDB atlas

---
## Instructions to use the application
1. install python
2. virtual env
  `python3 -m venv env`
3. activate virtual env `env\Scripts\activate.bat
(for windows)` `env/bin/activate
(for Unix or MacOS)` 
4. install pandas`pip install pandas` 
5. install flask `pip install flask`
6. install numpy `pip install numpy` 
7. install openpyxl `pip install openpyxl` 
8. install pymongo `pip install pymongo`
9. If you plan to use a local instance of MongoDB database, install MongoDB atlas. 
    * Create a MongoDB Cloud account.
    * Create a MongoDB Atlas cluster.
    * Configure network access and create a cluster user.
    * Connect to the cluster.
10. In the local project directory, create a new file called ".env".
11. Setup the environment variables as described above.

---
## Inputs required
- Excel sheet- A file containing scholar number and name of each student of a batch.
- CSV file- Downloaded attendance file from Microsoft Teams

### Csv File


