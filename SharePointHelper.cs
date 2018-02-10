using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.Taxonomy;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SGMLExtracter
{
    internal class SharePointHelper
    {

        internal const int LCID = 1033;
        private ClientContext _ctx = null;

        public SharePointHelper(string siteurl)
        {
            _ctx = new ClientContext(siteurl);
        }
        public SharePointHelper(string siteurl, System.Net.NetworkCredential cred)
        {
            _ctx = new ClientContext(siteurl);
            _ctx.Credentials = cred;
        }
        internal ClientContext Ctx
        {
            get
            { return _ctx; }
        }
        public void UploadDocument(string siteURL, string documentListName, string documentListURL, string documentName, byte[] documentStream)
        {

            using (ClientContext clientContext = new ClientContext(siteURL))
            {

                //Get Document List
                List documentsList = clientContext.Web.Lists.GetByTitle(documentListName);

                var fileCreationInformation = new FileCreationInformation();
                //Assign to content byte[] i.e. documentStream

                fileCreationInformation.Content = documentStream;
                //Allow owerwrite of document

                fileCreationInformation.Overwrite = true;
                //Upload URL

                fileCreationInformation.Url = siteURL + documentListURL + documentName;
                Microsoft.SharePoint.Client.File uploadFile = documentsList.RootFolder.Files.Add(
                    fileCreationInformation);

                //Update the metadata for a field having name "DocType"
                uploadFile.ListItemAllFields["DocType"] = "Favourites";

                uploadFile.ListItemAllFields.Update();
                clientContext.ExecuteQuery();

            }
        }

        public List CreateDocumentLibrary(string DocLibName, string DocDesc)
        {
            Web web = Ctx.Web;
            ListCreationInformation creationInfo = new ListCreationInformation();
            creationInfo.Title = DocLibName;
            creationInfo.TemplateType = (int)ListTemplateType.DocumentLibrary;
            List list = web.Lists.Add(creationInfo);
            list.Description = DocDesc;
            list.Update();
            Ctx.ExecuteQuery();
            return list;
        }
        public void DeleteDocumentLibrary(string DocLibName)
        {
            try
            {
                Web web = Ctx.Web;
                List oList = web.Lists.GetByTitle(DocLibName);

                if (oList != null)
                {
                    oList.DeleteObject();
                    Ctx.ExecuteQuery();
                }
            }
            catch (Exception ex)
            {
                string foo = ex + ""; //TODO
            }
        }
        public TermGroup CreateGroup(TermStore termStore, string groupname)
        {
            TermGroup group = termStore.Groups.FirstOrDefault(x => x.Name == groupname);
            if (group == null)
            {
                group = termStore.CreateGroup(groupname, Guid.NewGuid());

            }
            Ctx.Load(group);
            Ctx.ExecuteQuery();
            return group;
        }
        public TermSet CreateTermSet(TermGroup group, string termsetname)
        {
            Ctx.Load(group.TermSets);
            Ctx.ExecuteQuery();
            TermSet set = group.TermSets.FirstOrDefault(x => x.Name == termsetname);
            if (set == null)
            {
                set = group.CreateTermSet(termsetname, Guid.NewGuid(), LCID);
                Ctx.Load(set);
                Ctx.ExecuteQuery();
            }
            return set;
        }
        public TermStore GetTermStore()
        {
            TaxonomySession taxonomySession = TaxonomySession.GetTaxonomySession(Ctx);
            taxonomySession.UpdateCache();
            TermStore termStore = taxonomySession.GetDefaultSiteCollectionTermStore();
            Ctx.Load(termStore,
                termStoreArg => termStoreArg.WorkingLanguage,
                termStoreArg => termStoreArg.Id,
                termStoreArg => termStoreArg.Groups.Include(
                    groupArg => groupArg.Id,
                    groupArg => groupArg.Name
                )
            );
            Ctx.ExecuteQuery();
            return termStore;
        }
        public void CreateTerm(TermStore termStore, string Group, string TermSetName, string termset)
        {
            // Get the term group by Name
            TermGroup termGroup = termStore.Groups.GetByName(Group);
            // Get the term set by Name
            TermSet termSet = termGroup.TermSets.GetByName(TermSetName);
            int lcid = 1033;
            Term newTerm = termSet.CreateTerm(termset, lcid, Guid.NewGuid());
            Ctx.ExecuteQuery();
        }

        public void AddTaxFieldToList(List list, Field field, bool AddDefaultView)
        {
            list.Fields.Add(field);
            Ctx.ExecuteQuery();
        }

        internal void DeleteFieldIfExists(string fieldName)
        {
            Web web = Ctx.Web;
            Ctx.Load(web.Fields);
            Ctx.ExecuteQuery();
            Field field = web.Fields.FirstOrDefault(x => x.StaticName == fieldName);
            field.DeleteObject();
            Ctx.Load(field);
            Ctx.ExecuteQuery();
        }

        //internal  Field CreateTaxonomyField(TermStore termStore, TermSet termSet, string fieldName, bool isMulti, bool required)
        //{
        //    Web web = Ctx.Web;
        //    Guid txtFieldId = Guid.NewGuid();
        //    Guid taxFieldId = Guid.NewGuid();
        //    //Single valued, or multiple choice?
        //    string txType = isMulti ? "TaxonomyFieldTypeMulti" : "TaxonomyFieldType";
        //    //If it's single value, index it.
        //    string mult = isMulti ? "Mult='TRUE'" : "Indexed='TRUE'";
        //    string taxField = string.Format("<Field Type='{0}' DisplayName='{1}' ID='{8}' ShowField='Term1033' Required='{2}' EnforceUniqueValues='FALSE' {3} Sortable='FALSE' Name='{4}' Group='My Group'><Default/><Customization><ArrayOfProperty><Property><Name>SspId</Name><Value xmlns:q1='http://www.w3.org/2001/XMLSchema' p4:type='q1:string' xmlns:p4='http://www.w3.org/2001/XMLSchema-instance'>{5}</Value></Property><Property><Name>GroupId</Name></Property><Property><Name>TermSetId</Name><Value xmlns:q2='http://www.w3.org/2001/XMLSchema' p4:type='q2:string' xmlns:p4='http://www.w3.org/2001/XMLSchema-instance'>{6}</Value></Property><Property><Name>AnchorId</Name><Value xmlns:q3='http://www.w3.org/2001/XMLSchema' p4:type='q3:string' xmlns:p4='http://www.w3.org/2001/XMLSchema-instance'>00000000-0000-0000-0000-000000000000</Value></Property><Property><Name>UserCreated</Name><Value xmlns:q4='http://www.w3.org/2001/XMLSchema' p4:type='q4:boolean' xmlns:p4='http://www.w3.org/2001/XMLSchema-instance'>false</Value></Property><Property><Name>Open</Name><Value xmlns:q5='http://www.w3.org/2001/XMLSchema' p4:type='q5:boolean' xmlns:p4='http://www.w3.org/2001/XMLSchema-instance'>false</Value></Property><Property><Name>TextField</Name><Value xmlns:q6='http://www.w3.org/2001/XMLSchema' p4:type='q6:string' xmlns:p4='http://www.w3.org/2001/XMLSchema-instance'>{7}</Value></Property><Property><Name>IsPathRendered</Name><Value xmlns:q7='http://www.w3.org/2001/XMLSchema' p4:type='q7:boolean' xmlns:p4='http://www.w3.org/2001/XMLSchema-instance'>true</Value></Property><Property><Name>IsKeyword</Name><Value xmlns:q8='http://www.w3.org/2001/XMLSchema' p4:type='q8:boolean' xmlns:p4='http://www.w3.org/2001/XMLSchema-instance'>false</Value></Property><Property><Name>TargetTemplate</Name></Property><Property><Name>CreateValuesInEditForm</Name><Value xmlns:q9='http://www.w3.org/2001/XMLSchema' p4:type='q9:boolean' xmlns:p4='http://www.w3.org/2001/XMLSchema-instance'>false</Value></Property><Property><Name>FilterAssemblyStrongName</Name><Value xmlns:q10='http://www.w3.org/2001/XMLSchema' p4:type='q10:string' xmlns:p4='http://www.w3.org/2001/XMLSchema-instance'>Microsoft.SharePoint.Taxonomy, Version=16.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c</Value></Property><Property><Name>FilterClassName</Name><Value xmlns:q11='http://www.w3.org/2001/XMLSchema' p4:type='q11:string' xmlns:p4='http://www.w3.org/2001/XMLSchema-instance'>Microsoft.SharePoint.Taxonomy.TaxonomyField</Value></Property><Property><Name>FilterMethodName</Name><Value xmlns:q12='http://www.w3.org/2001/XMLSchema' p4:type='q12:string' xmlns:p4='http://www.w3.org/2001/XMLSchema-instance'>GetFilteringHtml</Value></Property><Property><Name>FilterJavascriptProperty</Name><Value xmlns:q13='http://www.w3.org/2001/XMLSchema' p4:type='q13:string' xmlns:p4='http://www.w3.org/2001/XMLSchema-instance'>FilteringJavascript</Value></Property></ArrayOfProperty></Customization></Field>",
        //    txType,fieldName,required.ToString().ToUpper(),mult,fieldName.Replace(" ", ""),termStore.Id.ToString("D"),
        //termSet.Id.ToString("D"),txtFieldId.ToString("B"),taxFieldId.ToString("B"));

        //    Field f = web.Fields.AddFieldAsXml(taxField, true, AddFieldOptions.AddFieldInternalNameHint);
        //    Ctx.Load(f);
        //    Ctx.ExecuteQuery();
        //    return f;
        //}
        internal TaxonomyField CreateTaxonomyField(TermStore termStore, TermSet termSet, string fieldName,string displayName, bool isMulti, bool required)
        {
            Web web = Ctx.Web;
            string columnTaxonomySchema = String.Format("<Field Type='TaxonomyFieldType' Name='{0}' DisplayName='{1}' ShowField='Term1033'   />", fieldName, displayName);
            var store = GetTermStore();
            var taxField = web.Fields.AddFieldAsXml(columnTaxonomySchema, false, AddFieldOptions.DefaultValue);
            Ctx.Load(taxField);
            Ctx.ExecuteQuery();

            var taxfield2 = Ctx.CastTo<TaxonomyField>(taxField);
            taxfield2.SspId = store.Id;
            taxfield2.TermSetId = termSet.Id; ;
            taxfield2.Update();
            Ctx.ExecuteQuery();
            return taxfield2;

        }
    }
}
