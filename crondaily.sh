# Crontab example — daily at 6:00 (adjust time as needed):
#   0 6 * * * /bin/bash /home/pi/work/acccnEExp/crondaily.sh
#
cd  /home/pi/work/acccnEExp
/usr/bin/node /home/pi/work/acccnEExp/dist/test.js sendDaily >> /home/pi/work/logs/acccneexp.log
