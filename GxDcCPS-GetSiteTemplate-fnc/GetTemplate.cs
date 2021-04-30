using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Threading.Tasks;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.Http;
using Microsoft.Azure.WebJobs.Host;
using Microsoft.SharePoint.Client;

namespace GxDcCPSGetSiteTemplatefnc
{
    public static class GetTemplate
    {
        [FunctionName("GetTemplate")]
        public static async Task<HttpResponseMessage> Run([HttpTrigger(AuthorizationLevel.Anonymous, "get", "post", Route = null)]HttpRequestMessage req, TraceWriter log)
        {
            string siteURL = "https://tbssctdev.sharepoint.com/teams/scw";


            // parse query parameter  
            log.Info("C# HTTP trigger function processed a request.");


            // Get request body  
            dynamic data = await req.Content.ReadAsAsync<object>();


            // SharePoint App only     
            ClientContext ctx = new OfficeDevPnP.Core.AuthenticationManager().GetAppOnlyAuthenticatedContext(siteURL, "9b92c43f-3a7c-42f8-b8e6-46f302dd5062", "xnxrDaxo4C6ZYW/DUCkO0myF4FHEHs/NTS+vdRenK3Y=");

            Web web = ctx.Web;
            List list = ctx.Web.Lists.GetByTitle("Space templates");
            ctx.Load(list);
            ctx.ExecuteQuery();


            CamlQuery camlQuery = new CamlQuery();
            camlQuery.ViewXml = "<View><RowLimit>100</RowLimit></View>";

            ListItemCollection collListItem = list.GetItems(camlQuery);
            ctx.Load(collListItem);

            ctx.ExecuteQuery();

            List<Dictionary<string, object>> result = new List<Dictionary<string, object>>();

            foreach (ListItem oListItem in collListItem)
            {
                log.Info("result");
                result.Add(new Dictionary<string, object>()
                {
                    {"TitleEn",oListItem["Template_x0020_Name_x0020__x0028"]},
                    {"TitleFr",oListItem["Template_x0020_Name_x0020__x00280"]},
                    {"DescriptionEn",oListItem["Template_x0020_Description_x0020"]},
                    {"DescriptionFr",oListItem["Template_x0020_Description_x00200"]},
                    {"TemplateImgUrl",oListItem["Template_x0020_Image_x0020_URL0"]}
                });
            }

            req.CreateResponse(HttpStatusCode.OK, "Create item successfully ");
            HttpResponseMessage response = req.CreateResponse(HttpStatusCode.OK, result);
            return response;
        }
    }
}
