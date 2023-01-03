using Microsoft.Office.Interop.Outlook;
using OutlookApp = Microsoft.Office.Interop.Outlook.Application;

OutlookApp outlookApp = new OutlookApp();
MailItem mailItem = outlookApp.CreateItem(OlItemType.olMailItem);

mailItem.To = "your@email.com";
mailItem.Subject = $"Subject";
mailItem.HTMLBody = "";

mailItem.Display(false);
mailItem.Send();

Console.ReadLine();

// Opciones -> Centro de confianza -> Acceso mediante programación