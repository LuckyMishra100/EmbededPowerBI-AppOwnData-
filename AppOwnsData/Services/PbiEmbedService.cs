﻿// ----------------------------------------------------------------------------
// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
// ----------------------------------------------------------------------------

namespace AppOwnsData.Services
{
    using AppOwnsData.Models;
    using Microsoft.PowerBI.Api;
    using Microsoft.PowerBI.Api.Models;
    using Microsoft.Rest;
    using System;
    using System.Collections.Generic;
    using System.IO;
    using System.Linq;
    using System.Runtime.InteropServices;
    using System.Threading.Tasks;

    public class PbiEmbedService
    {
        private readonly AadService aadService;
        private readonly string powerBiApiUrl  = "https://api.powerbi.com";

        public PbiEmbedService(AadService aadService)
        {
            this.aadService = aadService;
        }

        /// <summary>
        /// Get Power BI client
        /// </summary>
        /// <returns>Power BI client object</returns>
        public PowerBIClient GetPowerBIClient()
        {
            var tokenCredentials = new TokenCredentials(aadService.GetAccessToken(), "Bearer");
            return new PowerBIClient(new Uri(powerBiApiUrl ), tokenCredentials);
        }

        /// <summary>
        /// Get embed params for a report
        /// </summary>
        /// <returns>Wrapper object containing Embed token, Embed URL, Report Id, and Report name for single report</returns>
        public async Task<EmbedParams> GetEmbedParamsAsync(Guid workspaceId, Guid reportId, [Optional] Guid additionalDatasetId)
        {
            PowerBIClient pbiClient = this.GetPowerBIClient();

            // Get report info
            var pbiReport = pbiClient.Reports.GetReportInGroup(workspaceId, reportId);
            
            //  Check if dataset is present for the corresponding report
            //  If isRDLReport is true then it is a RDL Report 
            var isRDLReport = String.IsNullOrEmpty(pbiReport.DatasetId);

            EmbedToken embedToken;

            // Generate embed token for RDL report if dataset is not present
            if (isRDLReport)
            {
                // Get Embed token for RDL Report
                embedToken = GetEmbedTokenForRDLReport(workspaceId, reportId);
            }
            else
            {
                // Create list of datasets
                var datasetIds = new List<Guid>();

                // Add dataset associated to the report
                datasetIds.Add(Guid.Parse(pbiReport.DatasetId));

                // Append additional dataset to the list to achieve dynamic binding later
                if (additionalDatasetId != Guid.Empty)
                {
                    datasetIds.Add(additionalDatasetId);
                }

                // Get Embed token multiple resources
                embedToken = GetEmbedToken(reportId, datasetIds, workspaceId);
            }

            // Add report data for embedding
            var embedReports = new List<EmbedReport>() {
                new EmbedReport
                {
                    ReportId = pbiReport.Id, ReportName = pbiReport.Name, EmbedUrl = pbiReport.EmbedUrl
                }
            };

            // Capture embed params
            var embedParams = new EmbedParams
            {
                EmbedReport = embedReports,
                Type = "Report",
                EmbedToken = embedToken
            };

            return embedParams;
        }

        /// <summary>
        /// Get embed params for multiple reports for a single workspace
        /// </summary>
        /// <returns>Wrapper object containing Embed token, Embed URL, Report Id, and Report name for multiple reports</returns>
        /// <remarks>This function is not supported for RDL Report</remakrs>
        public EmbedParams GetEmbedParams(Guid workspaceId, IList<Guid> reportIds, [Optional] IList<Guid> additionalDatasetIds)
        {
            // Note: This method is an example and is not consumed in this sample app

            PowerBIClient pbiClient = this.GetPowerBIClient();

            // Create mapping for reports and Embed URLs
            var embedReports = new List<EmbedReport>();

            // Create list of datasets
            var datasetIds = new List<Guid>();

            // Get datasets and Embed URLs for all the reports
            foreach (var reportId in reportIds)
            {
                // Get report info
                var pbiReport = pbiClient.Reports.GetReportInGroup(workspaceId, reportId);

                datasetIds.Add(Guid.Parse(pbiReport.DatasetId));

                // Add report data for embedding
                embedReports.Add(new EmbedReport { ReportId = pbiReport.Id, ReportName = pbiReport.Name, EmbedUrl = pbiReport.EmbedUrl });
            }

            // Append to existing list of datasets to achieve dynamic binding later
            if (additionalDatasetIds != null)
            {
                datasetIds.AddRange(additionalDatasetIds);
            }

            // Get Embed token multiple resources
            var embedToken = GetEmbedToken(reportIds, datasetIds, workspaceId);

            // Capture embed params
            var embedParams = new EmbedParams
            {
                EmbedReport = embedReports,
                Type = "Report",
                EmbedToken = embedToken
            };

            return embedParams;
        }

        /// <summary>
        /// Get Embed token for single report, multiple datasets, and an optional target workspace
        /// </summary>
        /// <returns>Embed token</returns>
        /// <remarks>This function is not supported for RDL Report</remakrs>
        public EmbedToken GetEmbedToken(Guid reportId, IList<Guid> datasetIds, [Optional] Guid targetWorkspaceId)
        {
            PowerBIClient pbiClient = this.GetPowerBIClient();

            // Create a request for getting Embed token 
            // This method works only with new Power BI V2 workspace experience
            var tokenRequest = new GenerateTokenRequestV2(

                reports: new List<GenerateTokenRequestV2Report>() { new GenerateTokenRequestV2Report(reportId) },

                datasets: datasetIds.Select(datasetId => new GenerateTokenRequestV2Dataset(datasetId.ToString())).ToList(),

                targetWorkspaces: targetWorkspaceId != Guid.Empty ? new List<GenerateTokenRequestV2TargetWorkspace>() { new GenerateTokenRequestV2TargetWorkspace(targetWorkspaceId) } : null
            );

            // Generate Embed token
            var embedToken = pbiClient.EmbedToken.GenerateToken(tokenRequest);

            return embedToken;
        }

        /// <summary>
        /// Get Embed token for multiple reports, datasets, and an optional target workspace
        /// </summary>
        /// <returns>Embed token</returns>
        /// <remarks>This function is not supported for RDL Report</remakrs>
        public EmbedToken GetEmbedToken(IList<Guid> reportIds, IList<Guid> datasetIds, [Optional] Guid targetWorkspaceId)
        {
            // Note: This method is an example and is not consumed in this sample app

            PowerBIClient pbiClient = this.GetPowerBIClient();

            // Convert report Ids to required types
            var reports = reportIds.Select(reportId => new GenerateTokenRequestV2Report(reportId)).ToList();

            // Convert dataset Ids to required types
            var datasets = datasetIds.Select(datasetId => new GenerateTokenRequestV2Dataset(datasetId.ToString())).ToList();

            // Create a request for getting Embed token 
            // This method works only with new Power BI V2 workspace experience
            var tokenRequest = new GenerateTokenRequestV2(

                datasets: datasets,

                reports: reports,

                targetWorkspaces: targetWorkspaceId != Guid.Empty ? new List<GenerateTokenRequestV2TargetWorkspace>() { new GenerateTokenRequestV2TargetWorkspace(targetWorkspaceId) } : null
            );

            // Generate Embed token
            var embedToken = pbiClient.EmbedToken.GenerateToken(tokenRequest);

            return embedToken;
        }

        /// <summary>
        /// Get Embed token for multiple reports, datasets, and optional target workspaces
        /// </summary>
        /// <returns>Embed token</returns>
        /// <remarks>This function is not supported for RDL Report</remakrs>
        public EmbedToken GetEmbedToken(IList<Guid> reportIds, IList<Guid> datasetIds, [Optional] IList<Guid> targetWorkspaceIds)
        {
            // Note: This method is an example and is not consumed in this sample app

            PowerBIClient pbiClient = this.GetPowerBIClient();

            // Convert report Ids to required types
            var reports = reportIds.Select(reportId => new GenerateTokenRequestV2Report(reportId)).ToList();

            // Convert dataset Ids to required types
            var datasets = datasetIds.Select(datasetId => new GenerateTokenRequestV2Dataset(datasetId.ToString())).ToList();

            // Convert target workspace Ids to required types
            IList<GenerateTokenRequestV2TargetWorkspace> targetWorkspaces = null;
            if (targetWorkspaceIds != null)
            {
                targetWorkspaces = targetWorkspaceIds.Select(targetWorkspaceId => new GenerateTokenRequestV2TargetWorkspace(targetWorkspaceId)).ToList();
            }

            // Create a request for getting Embed token 
            // This method works only with new Power BI V2 workspace experience
            var tokenRequest = new GenerateTokenRequestV2(

                datasets: datasets,

                reports: reports,

                targetWorkspaces: targetWorkspaceIds != null ? targetWorkspaces : null
            );

            // Generate Embed token
            var embedToken = pbiClient.EmbedToken.GenerateToken(tokenRequest);

            return embedToken;
        }

        /// <summary>
        /// Get Embed token for RDL Report
        /// </summary>
        /// <returns>Embed token</returns>
        public EmbedToken GetEmbedTokenForRDLReport(Guid targetWorkspaceId, Guid reportId, string accessLevel = "view")
        {
            PowerBIClient pbiClient = this.GetPowerBIClient();

            // Generate token request for RDL Report
            var generateTokenRequestParameters = new GenerateTokenRequest(
                accessLevel: accessLevel
            );

            // Generate Embed token
            var embedToken = pbiClient.Reports.GenerateTokenInGroup(targetWorkspaceId, reportId, generateTokenRequestParameters);

            return embedToken;
        }


        ////////////////////////////////////////////////////////////
        ///
        public async Task<string> ExportPaginatedReportAsync(Guid WorkspaceId, Guid ReportId, string ExportName, FileFormat ExportFileFormat, List<ParameterValue> Parameters =null, string OutputFormat = "pdf")
        {
            try
            {
                PowerBIClient pbiClient = this.GetPowerBIClient();
                Console.WriteLine("Exporting " + ExportName + " (paginated report) to " + ((ExportFileFormat.Equals(FileFormat.IMAGE)) ? OutputFormat.ToUpper() : ExportFileFormat.ToString()));

                var exportRequest = new ExportReportRequest
                {
                    Format = ExportFileFormat,
                    PaginatedReportConfiguration = new PaginatedReportExportConfiguration
                    {
                        ParameterValues = new List<ParameterValue>() 
                        { new ParameterValue { Name = "ProfitCentre", Value = "2865" } 
                        , new ParameterValue { Name = "Year", Value = "2023" }  
                        , new ParameterValue { Name = "Period", Value = "1" }
                        ,new ParameterValue { Name = "WeekEndDate", Value = "1/08/2023 12:00:00 AM" } 

                        },
                        FormatSettings = new Dictionary<string, string>() {
              { "OutputFormat", OutputFormat }
            }
                    }
                };


                if (Parameters != null)
                {
                    exportRequest.PaginatedReportConfiguration.ParameterValues = Parameters;
                }


                // var  export2 = await pbiClient.Reports.GetReportsInGroupAsync(WorkspaceId);
                var export = await pbiClient.Reports.ExportToFileInGroupAsync(WorkspaceId, ReportId, exportRequest);

                string exportId = export.Id;

                do
                {
                    System.Threading.Thread.Sleep(10000);
                    export = pbiClient.Reports.GetExportToFileStatusInGroup(WorkspaceId, ReportId, exportId);
                    Console.WriteLine(" - Export status: " + export.PercentComplete.ToString() + "% complete");
                } while (export.Status != ExportState.Succeeded && export.Status != ExportState.Failed);

                if (export.Status == ExportState.Failed)
                {
                    Console.WriteLine("Export failed!");
                }

                string FilePath = "";
                if (export.Status == ExportState.Succeeded)
                {
                    string FileName = ExportName + export.ResourceFileExtension.ToLower();
                    string path = "E:\\Lucky_Demo_Projects\\PowerBI-Developer-Samples-master\\.NET Core\\Embed for your customers\\AppOwnsData\\PDF\\";
                    FilePath = path + FileName;

                    Console.WriteLine(" - Saving exported file to " + FilePath);
                    var exportStream = pbiClient.Reports.GetFileOfExportToFileInGroup(WorkspaceId, ReportId, exportId);
                    FileStream fileStream = File.Create(FilePath);
                    exportStream.CopyTo(fileStream);
                    fileStream.Close();
                }
                return (export.Status.ToString() + "$" + FilePath);
            }
            catch (Exception ex)
            {
                return ("Exception $"+ex.Message.ToString());
            }
           

        }
    }  
}
