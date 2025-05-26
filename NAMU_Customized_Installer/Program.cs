// Copyright (c) 2025 World Health Organization
// SPDX-License-Identifier: BSD-3-Clause

using System;
using System.IO;
using System.Collections;
using System.ComponentModel;
using System.Configuration.Install;
using Microsoft.VisualStudio.Tools.Applications;
using Microsoft.VisualStudio.Tools.Applications.Runtime;

namespace AddCustomizationCustomAction
{
    [RunInstaller(true)]
    public class AddCustomizations : Installer
    {
        public AddCustomizations() : base() { }

        public override void Install(IDictionary savedState)
        {
            base.Install(savedState);

            //Get the CustomActionData Parameters
            string documentLocation = Context.Parameters.ContainsKey("documentLocation") ? Context.Parameters["documentLocation"] : String.Empty;
            string assemblyLocation = Context.Parameters.ContainsKey("assemblyLocation") ? Context.Parameters["assemblyLocation"] : String.Empty;
            string deploymentManifestLocation = Context.Parameters.ContainsKey("deploymentManifestLocation") ? Context.Parameters["deploymentManifestLocation"] : String.Empty;
            Guid solutionID = Context.Parameters.ContainsKey("solutionID") ? new Guid(Context.Parameters["solutionID"]) : new Guid();

            string newDocLocation = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), Path.GetFileName(documentLocation));

            try
            {
                //Move the file and set the Customizations
                if (Uri.TryCreate(deploymentManifestLocation, UriKind.Absolute, out Uri docManifestLocationUri))
                {
                    File.Move(documentLocation, newDocLocation);
                    ServerDocument.RemoveCustomization(newDocLocation);
                    ServerDocument.AddCustomization(newDocLocation, assemblyLocation,
                                                    solutionID, docManifestLocationUri,
                                                    true, out string[] nonpublicCachedDataMembers);
                }
                else
                {
                    LogMessage("The document could not be customized.");
                }
            }
            catch (ArgumentException)
            {
                LogMessage("The document could not be customized.");
            }
            catch (DocumentNotCustomizedException)
            {
                LogMessage("The document could not be customized.");
            }
            catch (InvalidOperationException)
            {
                LogMessage("The customization could not be removed.");
            }
            catch (IOException)
            {
                LogMessage("The document does not exist or is read-only.");
            }
        }

        public override void Rollback(IDictionary savedState)
        {
            base.Rollback(savedState);
            DeleteDocument();
        }
        public override void Uninstall(IDictionary savedState)
        {
            base.Uninstall(savedState);
            DeleteDocument();
        }
        private void DeleteDocument()
        {
            string documentLocation = Context.Parameters.ContainsKey("documentLocation") ? Context.Parameters["documentLocation"] : String.Empty;

            try
            {
                File.Delete(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), Path.GetFileName(documentLocation)));
            }
            catch (Exception)
            {
                LogMessage("The document doesn't exist or is read-only.");
            }
        }
        private void LogMessage(string Message)
        {
            if (Context.Parameters.ContainsKey("LogFile"))
            {
                Context.LogMessage(Message);
            }
        }

        static void Main() { }
    }
}