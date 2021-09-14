using System.Threading.Tasks;
using Microsoft.Exchange.WebServices.Data;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Net;
using System.Security;
using System;
using System.ComponentModel;
using System.Diagnostics;
using System.Security;

namespace ShanghaiTechMail
{
    class Program
    {
        static bool RedirectionCallback(string url)
        {
            return url.ToLower().StartsWith("https://");
        }
        private static bool RedirectionUrlValidationCallback(string redirectionUrl)
        {
            // The default for the validation callback is to reject the URL.
            bool result = false;
            Uri redirectionUri = new Uri(redirectionUrl);
            // Validate the contents of the redirection URL. In this simple validation
            // callback, the redirection URL is considered valid if it is using HTTPS
            // to encrypt the authentication credentials. 
            if (redirectionUri.Scheme == "https")
            {
                result = true;
            }
            return result;
        }
        public static ExchangeService UseExchangeService(string userEmailAddress, SecureString userPassword)
        {
            ExchangeService service = new ExchangeService(ExchangeVersion.Exchange2010_SP1);

            #region Authentication

            // Set specific credentials.
            service.Credentials = new NetworkCredential(userEmailAddress, userPassword);
            #endregion
            //ItemView view = new ItemView(int.MaxValue);
            //FindItemsResults<Item> findResults = service.FindItems(new[] { })
            //service.AutodiscoverUrl("qinfr@shanghaitech.edu.cn", RedirectionUrlValidationCallback);
            service.Url = new Uri("https://mail.shanghaitech.edu.cn/ews/Exchange.asmx");
  
            return service;
        }

        public static void BatchDeleteEmailItems(ExchangeService service, Collection<ItemId> itemIds)
        {
            // Delete the batch of email message objects.
            // This method call results in an DeleteItem call to EWS.
            ServiceResponseCollection<ServiceResponse> response = service.DeleteItems(itemIds, DeleteMode.SoftDelete, null, AffectedTaskOccurrence.AllOccurrences);

            // Check for success of the DeleteItems method call.
            // DeleteItems returns success even if it does not find all the item IDs.
            if (response.OverallResult == ServiceResult.Success)
            {
                Console.WriteLine("Email messages deleted successfully.\r\n");
            }
            // If the method did not return success, print a message.
            else
            {
                Console.WriteLine("Not all email messages deleted successfully.\r\n");
            }
        }

        public static Collection<EmailMessage> BatchGetEmailItems(ExchangeService service, Collection<ItemId> itemIds)
        {
            // Create a property set that limits the properties returned by the Bind method to only those that are required.
            PropertySet propSet = new PropertySet(BasePropertySet.IdOnly, EmailMessageSchema.Subject, EmailMessageSchema.ToRecipients);
            // Get the items from the server.
            // This method call results in a GetItem call to EWS.
            ServiceResponseCollection<GetItemResponse> response = service.BindToItems(itemIds, propSet);

            // Instantiate a collection of EmailMessage objects to populate from the values that are returned by the Exchange server.
            Collection<EmailMessage> messageItems = new Collection<EmailMessage>();
            foreach (GetItemResponse getItemResponse in response)
            {
                try
                {
                    Item item = getItemResponse.Item;
                    EmailMessage message = (EmailMessage)item;
                    messageItems.Add(message);
                    // Print out confirmation and the last eight characters of the item ID.
                    Console.WriteLine("Found item {0}.", message.Id.ToString().Substring(144));
                }
                catch (Exception ex)
                {
                    Console.WriteLine("Exception while getting a message: {0}", ex.Message);
                }
            }
            // Check for success of the BindToItems method call.
            if (response.OverallResult == ServiceResult.Success)
            {
                Console.WriteLine("All email messages retrieved successfully.");
                Console.WriteLine("\r\n");
            }
            return messageItems;
        }

        public static void DeleteSpecificFolder(ExchangeService service)
        {
            ItemView view = new ItemView(int.MaxValue);
            FindItemsResults<Item> findResults = service.FindItems(WellKnownFolderName.DeletedItems, view);
            //Instantiate collection of ItemIds
            Collection<ItemId> messageItems = new Collection<ItemId>();
            foreach (Item item in findResults.Items)
            {
                //Add ItemIds to collection
                messageItems.Add(item.Id);
            }
            foreach (Item item in findResults.Items)
            {
                Console.WriteLine(item.Subject);
            }
            var response = service.DeleteItems(messageItems, DeleteMode.SoftDelete, null, AffectedTaskOccurrence.AllOccurrences);
        }

        static void Main(string[] args)
        {
            Console.Write("Enter mail: ");
            var mail = Console.ReadLine();
            SecureString securePwd = new SecureString();
            Console.Write("Enter password: ");
            ConsoleKeyInfo key;
            do
            {
                key = Console.ReadKey(true);

                // Ignore any key out of range.
                if (((int)key.Key) >= 65 && ((int)key.Key <= 90))
                {
                    // Append the character to the password.
                    securePwd.AppendChar(key.KeyChar);
                    Console.Write("*");
                }
                // Exit if Enter key is pressed.
            } while (key.Key != ConsoleKey.Enter);
            Console.WriteLine();
            ExchangeService service = UseExchangeService(mail, securePwd);
            DeleteSpecificFolder(service);

        }
    }
}
