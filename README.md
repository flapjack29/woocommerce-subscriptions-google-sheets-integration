# woocommerce-google-sheets-integration for Subscriptions
This google sheet syncs subscription data from your WooCommerce store to Google sheets automatically. A free alternative to Zapier-WooCommerce subscription integration.

The sheet uses the WooCommerce Rest API v1 to connect to the WooCommerce store and sync the subscription data to a google sheet. The sheet will fetch subscription details like First Name, Last Name, Billing Address, Shipping address, Phone, Email, Price, Payment method, Items, Quantity  OrderId, Notes, Date, Refunds, Order key.

NOTE: Woocommerce REST API only supports website that has https enabled. If your website is not https, consider installing SSL certificate from https://www.freessl.com/ or your hosting provider. 

Below are the steps to set up the sheet:

# 1. Set up WooCommerce REST API in your WordPress website:

Steps can be found here: https://github.com/woocommerce/woocommerce/wiki/Getting-started-with-the-REST-API#generate-keys. You need to provide the API key and API secret generated from this step in the Google sheet.

# 2. Set up google sheets
If you know how to work with google app script, copy the code from .gs file to your script editor and set up your google sheet in the format of the template.

YOU SHOULD ALWAYS KEEP YOUR API CREDENTIALS as secret.

IMPORTANT: Once you get the edit permission, go to File > Make a copy and save the copy to your google drive. Now, this copy can be accessed only by you. Do not make any modifications / enter your API credentials in the original sheet as anyone can see the data in the public sheet.

# 3. Set up your google sheet

Enter your store URL (Should be in the format https://yourstore.com - notice that only https is supported), API Key, API secret, Date in the sheet and click on the Update button. The google sheets will ask for permission for the app to run. The sheet requires permission to 1. View and manage your spreadsheets in Google Drive, 2. Connect to an external service. It will show the app as unverified - This is because it is trying to access an external URL - the site you have given in the sheet in this case. Click on Advanced and proceed. Please feel free to take a look at the code in the script editor to see what the sheet is doing in case you are in doubt. 
If everything is set up correctly, the sheet will update with the subscription details from your woo-commerce store.


# 4. Set up automatic subscription sync 

If you want the sheet to update automatically, you can set up a trigger for the sheet. 
Click on Tools > Script Editor. A new window will pop up. 

On that page, click on Edit > Current Projects Trigger. In the pop-up, give the name 'subscription_sync' in the first field and time driven in the second field. You can set the update frequency also here. There is an option to update sheet daily, hourly, monthly and even on opening the sheet.

'subscription_sync' is the function that invokes the script that syncs data between the google sheets and your site. Setting a trigger will trigger this process periodically based on the preferences you set.

Contributors to this project: @mithunmanohar, @petrfaitl, @flapjack29

If you find this tool useful, consider buying the original creators a coffee! I just piggybacked off their work.

<a href="https://www.buymeacoffee.com/68smoil" target="_blank"><img src="https://www.buymeacoffee.com/assets/img/custom_images/orange_img.png" alt="Buy Me A Coffee" style="height: 41px !important;width: 174px !important;box-shadow: 0px 3px 2px 0px rgba(190, 190, 190, 0.5) !important;-webkit-box-shadow: 0px 3px 2px 0px rgba(190, 190, 190, 0.5) !important;" ></a>
