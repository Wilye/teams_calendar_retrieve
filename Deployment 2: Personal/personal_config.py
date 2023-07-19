from dotenv import load_dotenv
import os

load_dotenv('personal_config.env')  # take environment variables from .env.

client_id = os.getenv('CLIENT_ID')
client_secret = os.getenv('CLIENT_SECRET')
tenant_id = os.getenv('TENANT_ID')
group_id = os.getenv('GROUP_ID')
channel_id = os.getenv('CHANNEL_ID')