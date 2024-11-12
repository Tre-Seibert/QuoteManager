import requests
import time
from exchangelib import OAuth2Credentials, Configuration, Account, IMPERSONATION
import msal
import threading
import os
from dotenv import load_dotenv
from oauthlib.oauth2 import OAuth2Token 
from exchangelib.version import Version, EXCHANGE_O365

# Load environment variables from .env file
load_dotenv()

# Authentication

# Token refresh interval in seconds (1 hour)
TOKEN_REFRESH_INTERVAL = 60 * 60


# Set up MSAL client 
def initialize():
    # Read sensitive information from environment variables
    shared_mailbox = os.getenv("SHARED_MAILBOX")
    m_sClientID = os.getenv("MS_CLIENT_ID")
    m_sClientSecret = os.getenv("MS_CLIENT_SECRET")
    m_sTenant = os.getenv("MS_TENANT")
    authority = os.getenv("AUTHORITY")

    # Set up MSAL client for Client Credentials
    app = msal.ConfidentialClientApplication(
        client_id=m_sClientID,
        client_credential=m_sClientSecret,
        authority=authority
    )

    # Acquire the initial token
    access_token = acquire_token(app)

    # Set up exchangelib credentials with the token
    credentials = OAuth2Credentials(
        client_id=m_sClientID,
        client_secret=m_sClientSecret,
        tenant_id=m_sTenant,
        identity=None,
        access_token=access_token
    )

    # Create exchangelib configuration and account
    config = Configuration(
        server="outlook.office365.com",
        credentials=credentials,
        version=Version(build=EXCHANGE_O365)
    )
    account = Account(primary_smtp_address=shared_mailbox,
                      config=config, autodiscover=False, access_type=IMPERSONATION)

    return account, app, credentials


# Function to acquire or refresh token
def acquire_token(app):
    token_result = app.acquire_token_for_client(
        scopes=["https://outlook.office365.com/.default"])
    if "access_token" in token_result:
        oauth2_token = OAuth2Token(
            {'access_token': token_result['access_token']})
        return oauth2_token
    else:
        raise Exception("Token acquisition failed:",
                        token_result.get("error"))


# Function to get quote details (deliveryAmount, deliveryTax)


def get_quote_details(quote_id, api_key):
    # create url and headers
    url = f"https://api.kaseyaquotemanager.com/v1/quote/{quote_id}"
    headers = {
        "apiKey": api_key,
        "Content-Type": "application/json"
    }
    # send get request
    response = requests.get(url, headers=headers)

    # check and return reponse
    if response.status_code == 200:
        return response.json()
    else:
        print("Quote Details Error:", response.status_code, response.text)
        return None


# Function to get sales order lines for a specific sales order ID
# salesOrderId param is extracted from get_quote_details response
# (price, tax)


def get_sales_order_lines(sales_order_id, api_key, page=1, page_size=100, modified_after=None):
    # create url and headers
    url = "https://api.kaseyaquotemanager.com/v1/salesorderline"
    headers = {
        "apiKey": api_key,
        "Content-Type": "application/json"
    }

    # Create query parameters
    params = {
        "salesOrderID": sales_order_id,
        "page": page,
        "pageSize": page_size
    }

    # probs wont use -can filter based on modified date
    if modified_after:
        params["modifiedAfter"] = modified_after

    # send get request
    response = requests.get(url, headers=headers, params=params)

    # check and return response
    if response.status_code == 200:
        return response.json()
    else:
        print("Sales Order Lines Error:", response.status_code, response.text)
        return None


# Function to calculate the total amount
# Sums the total price, total tax, delivery amount, delivery tax


def calc_total(quote_details, sales_order_lines):
    total_price = sum(item['price'] * item.get('quantity', 1)
                      for item in sales_order_lines)
    total_tax = sum(item['tax'] * item.get('quantity', 1)
                    for item in sales_order_lines)
    delivery_amount = quote_details.get('deliveryAmount', 0)
    delivery_tax = quote_details.get('deliveryTax', 0)

    total = total_price + total_tax + delivery_amount + delivery_tax
    return total


# Function to return a list of titles and ids from the quote end point


def get_all_quotes(api_key, quote_number="", modified_after=None, page=1, page_size=100):
    all_quotes = []  # List to store all quotes
    quote_summary = []  # List to store id and title pairs

    while True:
        # Create URL and headers
        url = "https://api.kaseyaquotemanager.com/v1/quote"
        headers = {
            "apiKey": api_key,
            "Content-Type": "application/json"
        }

        # Create query parameters
        params = {
            "page": page,
            "pageSize": page_size,
            "quoteNumber": quote_number
        }

        if modified_after:
            params["modifiedAfter"] = modified_after

        # Send GET request
        response = requests.get(url, headers=headers, params=params)

        # Check and handle the response
        if response.status_code == 200:
            data = response.json()
            # Directly extend all_quotes with data if data is a list
            if isinstance(data, list):
                # Append current page's quotes to the list
                all_quotes.extend(data)
                # Extract id and title for each quote
                for quote in data:
                    quote_summary.append(
                        {'id': quote.get('id'), 'title': quote.get('title')})
            else:
                # Debugging information
                print("Unexpected response format:", data)

            # Check if there are more pages
            if len(data) < page_size:
                break  # Exit loop if the last page is reached
            page += 1  # Increment the page number for the next request
        else:
            print("Quotes Error:", response.status_code, response.text)
            break  # Exit the loop on error

    # Print the summary of quotes
    # print("Quote Summary (ID and Title):", quote_summary)
    return all_quotes


def monitor_inbox(api_key, account, app, credentials, check_interval=5):
    api_key = api_key

    # Get all quotes
    all_quotes = get_all_quotes(api_key)
    print(f"Retrieved {len(all_quotes)}")

    # Access "Processed Quotes" folder
    processed_quotes_folder = account.root // \
        "Top of Information Store" // "Processed Quotes"

    while True:
        print("Checking for new approved emails...")

        # Refresh the token to ensure it is valid
        access_token = acquire_token(app)

        # Reinitialize credentials with the new token
        new_credentials = OAuth2Credentials(
            client_id=credentials.client_id,
            client_secret=credentials.client_secret,
            tenant_id=credentials.tenant_id,
            identity=None,
            access_token=access_token
        )

        # Recreate account object with new credentials
        config = Configuration(
            server="outlook.office365.com",
            credentials=new_credentials,
            version=Version(build=EXCHANGE_O365)
        )
        account = Account(primary_smtp_address=account.primary_smtp_address,
                          config=config, autodiscover=False, access_type=IMPERSONATION)

        # Pull up to 1000 emails ordered by datetime received
        for item in account.inbox.all().order_by("-datetime_received")[:1000]:
            if "has been signed by " in item.subject:
                title = item.subject.split(" has been signed by ")[0]
                print(f"Email found with title: {title}")
                all_quotes = get_all_quotes(api_key)
                for quote in all_quotes:
                    if quote['title'] == title:
                        print("Title found in list of titles")
                        quoteId = quote['id']
                        quote_details = get_quote_details(quoteId, api_key)

                        if quote_details:
                            sales_order_id = quote_details.get('salesOrderId')

                            if sales_order_id:
                                sales_order_lines = get_sales_order_lines(
                                    sales_order_id, api_key)

                                if sales_order_lines:
                                    total_amount = calc_total(
                                        quote_details, sales_order_lines)
                                    print(f"The total amount for quote {quoteId} is: ${total_amount:.2f}")
                                    print(f"Moving email: {item.subject}")
                                    item.move(processed_quotes_folder)
                                    print(f"Email moved to: {processed_quotes_folder.name}")

        # Wait for the specified interval before checking again
        time.sleep(check_interval)


# Main Function
if __name__ == "__main__":
    api_key = os.getenv("API_KEY")
    # Initialize account and authentication
    account, app, credentials = initialize()

    # Start monitoring the inbox
    monitor_inbox(api_key, account, app, credentials, check_interval=5)
