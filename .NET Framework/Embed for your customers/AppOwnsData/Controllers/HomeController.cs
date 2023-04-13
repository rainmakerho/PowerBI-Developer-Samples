// ----------------------------------------------------------------------------
// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
// ----------------------------------------------------------------------------

namespace PowerBIEmbedded_AppOwnsData.Controllers
{
    using AppOwnsData.Models;
    using AppOwnsData.Services;
    using Microsoft.Rest;
    using System;
    using System.Configuration;
    using System.Linq;
    using System.Reflection;
    using System.Threading.Tasks;
    using System.Web.Mvc;
    using Microsoft.Ajax.Utilities;
    using System.Web;
    using System.Collections.Concurrent;

    public class HomeController : Controller
    {
        private string m_errorMessage;

        public HomeController()
        {
            m_errorMessage = ConfigValidatorService.GetWebConfigErrors();
        }

        public ActionResult Index()
        {
            // Assembly info is not needed in production apps and the following 6 lines can be removed

            var result = new IndexConfig();
            var assembly = Assembly.GetExecutingAssembly().GetReferencedAssemblies().Where(n => n.Name.Equals("Microsoft.PowerBI.Api")).FirstOrDefault();
            if (assembly != null)
            {
                result.DotNetSDK = assembly.Version.ToString(3);
            }
            return View(result);
        }

        public async Task<ActionResult> EmbedReport(string reportId)
        {
            if (!m_errorMessage.IsNullOrWhiteSpace())
            {
                return View("Error", BuildErrorModel(m_errorMessage));
            }

            try
            {
                if (string.IsNullOrWhiteSpace(reportId))
                {
                    reportId = ConfigValidatorService.ReportId.ToString();
                }
                //每次都取 Embed Config
                //var newRptEmbedConfig = await EmbedService.GetEmbedParams(ConfigValidatorService.WorkspaceId, Guid.Parse(reportId));
                //return View(newRptEmbedConfig);

                //有過期才取
                var rptEmbedConfigsName = "RptEmbedConfigs";
                var rptEmbedConfigs = HttpContext.Application[rptEmbedConfigsName] as ConcurrentDictionary<string, ReportEmbedConfig>;
                if (rptEmbedConfigs == null)
                {
                    int initialCapacity = 101;
                    int numProcs = Environment.ProcessorCount;
                    int concurrencyLevel = numProcs * 2;
                    rptEmbedConfigs = new ConcurrentDictionary<string, ReportEmbedConfig>(concurrencyLevel, initialCapacity);
                    HttpContext.Application.Lock();
                    HttpContext.Application[rptEmbedConfigsName] = rptEmbedConfigs;
                    HttpContext.Application.UnLock();
                }

                ReportEmbedConfig rptEmbedConfig = null;

                if (rptEmbedConfigs.TryGetValue(reportId, out rptEmbedConfig))
                {
                    if (rptEmbedConfig.EmbedToken.Expiration > DateTime.UtcNow.AddMinutes(5))
                    {
                        return View(rptEmbedConfig);
                    }
                }

                var newRptEmbedConfig = await EmbedService.GetEmbedParams(ConfigValidatorService.WorkspaceId, Guid.Parse(reportId));
                rptEmbedConfigs.TryAdd(reportId, newRptEmbedConfig);
                return View(newRptEmbedConfig);
            }
            catch (HttpOperationException exc)
            {
                m_errorMessage = string.Format("Status: {0} ({1})\r\nResponse: {2}\r\nRequestId: {3}", exc.Response.StatusCode, (int)exc.Response.StatusCode, exc.Response.Content, exc.Response.Headers["RequestId"].FirstOrDefault());
                return View("Error", BuildErrorModel(m_errorMessage));
            }
            catch (Exception ex)
            {
                return View("Error", BuildErrorModel(ex.Message));
            }
        }

        public async Task<ActionResult> EmbedDashboard()
        {
            if (!m_errorMessage.IsNullOrWhiteSpace())
            {
                return View("Error", BuildErrorModel(m_errorMessage));
            }

            try
            {
                var embedResult = await EmbedService.EmbedDashboard(new Guid(ConfigurationManager.AppSettings["workspaceId"]));
                return View(embedResult);
            }
            catch (HttpOperationException exc)
            {
                m_errorMessage = string.Format("Status: {0} ({1})\r\nResponse: {2}\r\nRequestId: {3}", exc.Response.StatusCode, (int)exc.Response.StatusCode, exc.Response.Content, exc.Response.Headers["RequestId"].FirstOrDefault());
                return View("Error", BuildErrorModel(m_errorMessage));
            }
            catch (Exception ex)
            {
                return View("Error", BuildErrorModel(ex.Message));
            }
        }

        public async Task<ActionResult> EmbedTile()
        {
            if (!m_errorMessage.IsNullOrWhiteSpace())
            {
                return View("Error", BuildErrorModel(m_errorMessage));
            }

            try
            {
                var embedResult = await EmbedService.EmbedTile(new Guid(ConfigurationManager.AppSettings["workspaceId"]));
                return View(embedResult);
            }
            catch (HttpOperationException exc)
            {
                m_errorMessage = string.Format("Status: {0} ({1})\r\nResponse: {2}\r\nRequestId: {3}", exc.Response.StatusCode, (int)exc.Response.StatusCode, exc.Response.Content, exc.Response.Headers["RequestId"].FirstOrDefault());
                return View("Error", BuildErrorModel(m_errorMessage));
            }
            catch (Exception ex)
            {
                return View("Error", BuildErrorModel(ex.Message));
            }
        }

        private ErrorModel BuildErrorModel(string errorMessage)
        {
            return new ErrorModel
            {
                ErrorMessage = errorMessage
            };
        }
    }
}
