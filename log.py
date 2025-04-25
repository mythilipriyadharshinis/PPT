import logging
from table3 import generate_ppt
from send_email import send_email

logging.basicConfig(
    filename='automation_log.log',     # Log file name
    level=logging.INFO,                # Log level: INFO, WARNING, ERROR
    format='%(asctime)s - %(levelname)s - %(message)s'
)
logging.info("Starting report generation...")

try:
    generate_ppt(master_client_id, booking_month, client_id=None, is_client_level=False)
except Exception as e:
    logging.error(f"Report generation failed for MasterID {master_id}, ClientID {client_id}: {e}")

try:
    send_email(master_id, client_id)
except Exception as e:
    logging.error(f"Email failed for MasterID {master_id}, ClientID {client_id}: {e}")
