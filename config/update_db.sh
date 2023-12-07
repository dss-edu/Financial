#!/bin/bash
source /var/www/Financial/config/venv/bin/activate
python /var/www/Financial/config/update_db.py
deactivate
