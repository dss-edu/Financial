#!/bin/bash
source /var/www/Financial/config/venv/bin/activate
python /var/www/Financial/config/update_db.py
deactivate
chown -R www-data:www-data /var/www/Financial/config/finance/json/
chmod -R 775 /var/www/Financial/config/finance/json/
