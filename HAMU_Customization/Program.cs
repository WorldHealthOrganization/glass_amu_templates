using Microsoft.VisualStudio.Tools.Applications;
using Microsoft.VisualStudio.Tools.Applications.Runtime;
using System;
using System.Collections;
using System.ComponentModel;
using System.Configuration.Install;
using System.IO;
using System.Runtime.Remoting.Contexts;

namespace AddHAMUCustomizationCustomAction
{
    [RunInstaller(true)]
    public class AddCustomizations : Installer
    {
        public AddCustomizations() : base() { }

        public override void Install(IDictionary savedState)
        {
            base.Install(savedState);

            //Get the CustomActionData Parameters
            //Command line custom action:
            // /assemblyLocation="[TARGETDIR]NAMU_Template.dll"
            // /deploymentManifestLocation="[TARGETDIR]NAMU_Template.vsto"
            // /documentLocation="[TARGETDIR]WHO_NAMU_Template.xlsx"
            // /solutionID="5a15c387-cb91-4590-bbf4-bf6ac3dcde2b"
            // /LogFile="[TARGETDIR]Setup.log"
            string documentLocation = Context.Parameters.ContainsKey("documentLocation") ? Context.Parameters["documentLocation"] : String.Empty;
            string assemblyLocation = Context.Parameters.ContainsKey("assemblyLocation") ? Context.Parameters["assemblyLocation"] : String.Empty;
            string deploymentManifestLocation = Context.Parameters.ContainsKey("deploymentManifestLocation") ? Context.Parameters["deploymentManifestLocation"] : String.Empty;
            Guid solutionID = Context.Parameters.ContainsKey("solutionID") ? new Guid(Context.Parameters["solutionID"]) : new Guid();

            //string newDocLocation = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), Path.GetFileName(documentLocation));

            try
            {
                //Set the Customizations
                if (Uri.TryCreate(deploymentManifestLocation, UriKind.Absolute, out Uri docManifestLocationUri))
                {
                    // File.Move(documentLocation, newDocLocation);
                    ServerDocument.RemoveCustomization(documentLocation);
                    ServerDocument.AddCustomization(documentLocation, assemblyLocation,
                                                    solutionID, docManifestLocationUri,
                                                    true, out string[] nonpublicCachedDataMembers);
                    LogMessage($"The document {documentLocation} has been customized.");
                }
                else
                {
                    LogMessage($"The URI from {deploymentManifestLocation} could not be created.");
                    LogMessage($"The document {documentLocation} could not be customized.");
                }
            }
            catch (ArgumentException e)
            {
                LogMessage("Exception during customization:");
                LogMessage(e.ToString());
                LogMessage($"The document {documentLocation} could not be customized.");
            }
            catch (DocumentNotCustomizedException e)
            {
                LogMessage("Exception during customization:");
                LogMessage(e.ToString());
                LogMessage($"The document {documentLocation} could not be customized.");
            }
            catch (InvalidOperationException e)
            {
                LogMessage("Exception during customization:");
                LogMessage(e.ToString());
                LogMessage("The customization could not be removed.");
            }
            catch (IOException e)
            {
                LogMessage("Exception during customization:");
                LogMessage(e.ToString());
                LogMessage($"The document {documentLocation} does not exist or is read-only.");
            }
        }

        public override void Rollback(IDictionary savedState)
        {
            DeleteLog();
            base.Rollback(savedState);
        }

        public override void Uninstall(IDictionary savedState)
        {
            DeleteLog();
            base.Uninstall(savedState);
        }

        //private void DeleteDocument()
        //{
        //    string documentLocation = Context.Parameters.ContainsKey("documentLocation") ? Context.Parameters["documentLocation"] : String.Empty;

        //    try
        //    {
        //        File.Delete(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), Path.GetFileName(documentLocation)));
        //    }
        //    catch (Exception)
        //    {
        //        LogMessage("The document doesn't exist or is read-only.");
        //    }
        //}

        private void LogMessage(string Message)
        {
            if (Context.Parameters.ContainsKey("LogFile"))
            {
                Context.LogMessage($"{DateTime.Now:yyyy-MM-dd HH:mm:ss}\t{Message}");
            }
        }

        private void DeleteLog()
        {
            if (Context.Parameters.ContainsKey("LogFile"))
            {
                string LogLocation = Context.Parameters["LogFile"];
                try
                {
                    File.Delete(LogLocation);
                }
                catch (Exception)
                {

                }
            }
        }

        static void Main() { }
    }
}
