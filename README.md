# Financial
Install 
  sudo apt install unixodbc
  sudo apt install msodbcsql17

# Configuring Database Update Schedule
### Date/Time
The database update is scheduled using `cron`. 

To configure the time or day of update. First, login as root user. Then, configure the cron scheduler using crontab.

```sh
sudo su # login as root user
crontab -e # edit cron scheduler
```
#### Database
The scheduler runs the `update_db.sh` script and the said script runs the `update_db.py`. 

`update_db.py` is the file you want to modify if you want to change which databases you want to update. The codes for the updaters are found in `backend.py`. 

As defulat, `update_db.py` uses the `update_fy()` function as the updater. This updates the data for each school in the **current** school year. 
