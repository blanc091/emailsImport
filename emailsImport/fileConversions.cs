using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
//css_ref GdPicture.NET.14.dll
//using GdPicture14;
using System.Windows.Forms;
using System.Drawing;
using static emailsImport.emImport;

namespace emailsImport
{
    internal class fileConversions
    {
        public void createInternalTIFFS(string pdfFullPath, string filePathTIFF, string folderPathGUID)
        {/*
            bool convertPDF = false;
            bool convertTIFFS = false;
            GdPicture14.LicenseManager oLicenseManager = new GdPicture14.LicenseManager();
            oLicenseManager.RegisterKEY("******");
            if (Path.GetExtension(pdfFullPath).ToLower() == ".pdf")
            {
                convertPDF = true;
            }
            else if (Path.GetExtension(pdfFullPath).ToLower() == ".tif" || Path.GetExtension(pdfFullPath).ToLower() == ".tiff")
            {
                convertTIFFS = true;
            }
            if (convertPDF)
            {
                GdPictureImaging oGdPictureImaging = new GdPictureImaging();
                GdPicturePDF oGdPicturePDF = new GdPicturePDF();  
                GdPictureStatus status = oGdPicturePDF.LoadFromFile(pdfFullPath, false);
                if (status != GdPictureStatus.OK)
                {
                    MessageBox.Show("The PDF document can't be loaded. Error: " + status.ToString(), "TIFF Conversions", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                else
                {
                    for (int i = 1; i <= oGdPicturePDF.GetPageCount(); i++)
                    {
                        oGdPicturePDF.SelectPage(i);
                        int imageId = oGdPicturePDF.RenderPageToGdPictureImage(300, true);
                        if (oGdPicturePDF.GetStat() == GdPictureStatus.OK)
                        {
                            string tiffFilePath = Path.Combine(filePathTIFF, $"{i}.tif");
                            oGdPictureImaging.TiffSaveAsMultiPageFile(imageId, tiffFilePath, TiffCompression.TiffCompressionLZW);
                            oGdPictureImaging.ReleaseGdPictureImage(imageId);
                            Logger.Log(GlobalConfig.Client, $"Page {i} saved as {tiffFilePath}");
                        }
                        else
                        {
                            Logger.Log(GlobalConfig.Client, $"Error occurred when rendering page {i} to an image. Error: {oGdPicturePDF.GetStat()}");
                        }
                    }
                    createBatch batch_create = new createBatch();
                    batch_create.createKofaxBatch(filePathTIFF, folderPathGUID);
                }
                oGdPictureImaging.Dispose();
            }
            else if (convertTIFFS)
            {
                using (GdPictureImaging oGdPictureImaging = new GdPictureImaging())
                {
                    oGdPictureImaging.TiffOpenMultiPageForWrite(true);
                    int ImageID = oGdPictureImaging.CreateGdPictureImageFromFile(pdfFullPath, true);
                    if (oGdPictureImaging.GetStat() == GdPictureStatus.OK)
                    {
                        int PageCount = oGdPictureImaging.GetPageCount(ImageID);
                        GdPictureStatus status = GdPictureStatus.OK;
                        if (PageCount > 0)
                        {
                            TiffCompression ImgTiffCompression = oGdPictureImaging.GetTiffCompression(ImageID);
                            for (int i = 1; i <= PageCount; i++)
                            {
                                status = oGdPictureImaging.TiffSelectPage(ImageID, i);
                                if (status == GdPictureStatus.OK)
                                {
                                   if (status == GdPictureStatus.OK)
                                        status = oGdPictureImaging.AppendToTiff(ImageID, i.ToString(), ImgTiffCompression);
                                    if (status != GdPictureStatus.OK) break;
                                }
                                else break;
                            }
                        }                        
                        if (status == GdPictureStatus.OK)
                            MessageBox.Show("The process has finished successfully.", "Append To TIFF Example", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        else
                            MessageBox.Show("The process has failed. Status: " + status.ToString(), "Append To TIFF Example", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        oGdPictureImaging.ReleaseGdPictureImage(ImageID);
                    }
                    else MessageBox.Show("The source file can't be loaded. Status: " + oGdPictureImaging.GetStat().ToString(), "Append To TIFF Example", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            */
        }             
    }
}
