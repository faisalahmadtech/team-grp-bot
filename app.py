import os
import json
import requests
from flask import Flask, request
from msal import ConfidentialClientApplication
from dotenv import load_dotenv

# Load environment variables
load_dotenv()


app = Flask(__name__)
@app.route("/", methods=["GET"])
def home():
    return "Bot is running!", 200

# Azure AD credentials
AZURE_APP_ID = os.getenv("AZURE_APP_ID")
AZURE_CLIENT_SECRET = os.getenv("AZURE_CLIENT_SECRET")
AZURE_TENANT_ID = os.getenv("AZURE_TENANT_ID")
OLLAMA_API_URL = os.getenv("OLLAMA_API_URL")

# Microsoft Graph API base URL
GRAPH_API_URL = "https://graph.microsoft.com/v1.0"

# MSAL app for token acquisition
msal_app = ConfidentialClientApplication(
    AZURE_APP_ID,
    authority=f"https://login.microsoftonline.com/{AZURE_TENANT_ID}",
    client_credential=AZURE_CLIENT_SECRET
)

def get_access_token():
    """Get access token for Microsoft Graph API."""
    result = msal_app.acquire_token_for_client(scopes=["https://graph.microsoft.com/.default"])
    if "access_token" in result:
        return result["access_token"]
    raise Exception("Failed to acquire access token.")

def send_message_to_channel(team_id, channel_id, message):
    """Send a message to a Microsoft Teams channel."""
    token = get_access_token()
    url = f"{GRAPH_API_URL}/teams/{team_id}/channels/{channel_id}/messages"
    headers = {
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/json"
    }
    payload = {
        "body": {
            "content": message
        }
    }
    response = requests.post(url, headers=headers, json=payload)
    if response.status_code != 201:
        raise Exception(f"Failed to send message: {response.text}")

def query_ollama_model(prompt):
    """Query the Ollama model for a response."""
    payload = {
        "model": "llama2",  # Replace with your desired model
        "prompt": prompt
    }
    response = requests.post(OLLAMA_API_URL, json=payload)
    if response.status_code == 200:
        return response.json().get("response", "No response from model.")
    raise Exception(f"Ollama API error: {response.text}")

@app.route("/api/messages", methods=["POST"])
def handle_message():
    """Handle incoming messages from Microsoft Teams."""
    data = request.json
    print("Received message:", data)

    # Extract message details
    team_id = data.get("teamId")
    channel_id = data.get("channelId")
    user_message = data.get("text", "").strip()

    if not user_message:
        return "No message content", 200

    try:
        # Query Ollama model for a response
        bot_response = query_ollama_model(user_message)

        # Send the bot's response back to the channel
        send_message_to_channel(team_id, channel_id, bot_response)
        return "Message processed", 200
    except Exception as e:
        print("Error:", str(e))
        return "Error processing message", 500

if __name__ == "__main__":
    app.run(port=5000)