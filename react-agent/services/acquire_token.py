import urllib.parse

import requests


class AcquireToken:
    """Acquire a token using MSAL."""

    def __init__(
        self,
        site_url,
        tenant_id,
        client_id,
        client_secret,
        resource_url,
        user_assertion=None,
    ):
        self.site_url = site_url
        self.tenant_id = tenant_id
        self.client_id = client_id
        self.client_secret = client_secret
        self.resource_url = resource_url
        self.base_url = (
            f"https://login.microsoftonline.com/{tenant_id}/oauth2/v2.0/token"
        )
        self.headers = {"Content-Type": "application/x-www-form-urlencoded"}
        self.user_assertion = user_assertion
        self.access_token = self.get_access_token()
        self.obo_access_token = self.get_obo_access_token()

    def get_access_token(self):
        """
        This function retrieves an access token from Microsoft's OAuth2 endpoint.

        The access token is used to authenticate and authorize the application for
        accessing Microsoft Graph API resources.

        Returns:
        str: The access token as a string. This token is used for authentication in subsequent API requests.
        """
        # Body for the access token request
        body = {
            "grant_type": "client_credentials",
            "client_id": self.client_id,
            "client_secret": self.client_secret,
            "scope": self.resource_url + ".default",
        }
        response = requests.post(self.base_url, headers=self.headers, data=body)
        return response.json().get(
            "access_token"
        )  # Extract access token from the response

    def get_obo_access_token(self):
        """
        This function retrieves an On-Behalf-Of (OBO) access token from Microsoft's OAuth2 endpoint.

        The OBO access token is used to authenticate and authorize the application to act on behalf of a user
        for accessing Microsoft Graph API resources.

        Args:
            user_assertion (str): The user's access token that the application is acting on behalf of.
        """
        # Body for the access token request
        body = {
            "client_id": self.client_id,
            "client_secret": self.client_secret,
            "grant_type": "urn:ietf:params:oauth:grant-type:jwt-bearer",
            "requested_token_use": "on_behalf_of",
            "assertion": self.user_assertion,
            "scope": self.resource_url + ".default",
        }
        response = requests.post(self.base_url, headers=self.headers, data=body)
        return response.json().get(
            "access_token"
        )  # Extract access token from the response
