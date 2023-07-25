from dotenv import load_dotenv
import os

dotenv_path = os.path.join(os.path.dirname(__file__), 'JERA_config.env')
load_dotenv(dotenv_path)

client_id = os.getenv('CLIENT_ID')
client_secret = os.getenv('CLIENT_SECRET')
tenant_id = os.getenv('TENANT_ID')
group_id = os.getenv('GROUP_ID')
channel_id = os.getenv('CHANNEL_ID')
