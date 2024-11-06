using System;
using System.IO;
using System.Runtime.InteropServices;

namespace emailsImport
{
    internal class createBatch
    {
        public void createKofaxBatch(string importFilesPath, string folderPathGUID)
        {/*
            //importFilesPath = @"C:\CTERA\CaptureSV\Test\";
            Kofax.AscentCaptureModule.ImportLogin myLogin;
            Kofax.AscentCaptureModule.Application myApp;

            Logger.Log(project, "Logging in into Kofax Capture..");
            myLogin = new Kofax.AscentCaptureModule.ImportLogin();
            myApp = myLogin.Login("********", "*******");
            Logger.Log(project, "Logged in.");
            Logger.Log(project, "Creating new batch..");
            Kofax.AscentCaptureModule.BatchClass myBatchClass = myApp.BatchClasses["EmailsImportTest"];
            Kofax.AscentCaptureModule.Batch myBatch = myApp.CreateBatch(ref myBatchClass, "EmailsImportTest");
            string[] folderPaths = Directory.GetDirectories(folderPathGUID);
            Logger.Log(project, "Creating a new document..");
            foreach (string folderGUID in folderPaths)
            {
                try
                {
                    Kofax.AscentCaptureModule.Document myDoc;
                    myDoc = myBatch.CreateDocument(null);
                    Logger.Log(project, "Assigning a form type..");
                    Kofax.AscentCaptureModule.FormType myFormType = myBatch.FormTypes[1];
                    myDoc.set_FormType(ref myFormType);
                    Logger.Log(project, "Adding pages to the document..");
                    string batchImagesFolderPath = Path.Combine(folderGUID, "Batch_Images");
                    string[] filePaths = Directory.GetFiles(batchImagesFolderPath);
                    foreach (string filePath in filePaths)
                    {
                        Kofax.AscentCaptureModule.Pages myPages = myBatch.ImportFile(filePath);
                        foreach (Kofax.AscentCaptureModule.Page myPage in myPages)
                        {
                            myPage.MoveToDocument(ref myDoc, null);
                            Logger.Log(project, "Added page..");
                        }
                    }
                }
                catch (Exception ex)
                {
                    Logger.Log(project, $"Error: {ex.Message}");
                }
            }
            myApp.CloseBatch();*/
        }
        }
}
