import urllib.parse

import requests


class SiteInfo:
    def __init__(self, site_url: str, access_token: str):
        self.site_url = site_url
        self.access_token = access_token
        self.site_id = self.get_site_id()

    def get_site_id(self):
        """
        This function retrieves the ID of a SharePoint site using the Microsoft Graph API.

        Returns:
        str: The ID of the SharePoint site.
        """
        # Parse the site URL
        parsed_url = urllib.parse.urlparse(self.site_url)
        hostname = parsed_url.hostname  # e.g., tenant.sharepoint.com
        # Extract site path relative to the root, e.g., 'sites/sitename'
        site_path = parsed_url.path.strip("/")

        # Construct the Graph API URL
        full_url = f"https://graph.microsoft.com/v1.0/sites/{hostname}:/{site_path}"

        headers = {"Authorization": f"Bearer {self.access_token}"}
        response = requests.get(full_url, headers=headers)
        response.raise_for_status()  # Raise error if request failed

        site_id = response.json().get("id")
        return site_id
