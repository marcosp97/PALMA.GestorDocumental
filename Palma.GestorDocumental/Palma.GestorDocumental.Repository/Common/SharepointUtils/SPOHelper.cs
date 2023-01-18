using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using SP = Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.Search.Query;
using Microsoft.SharePoint.Client.Utilities;
using System.Security;
using obj = Palma.GestorDocumental.Repository.Common.Configuration.Manager.ObjectSharepoint.Site;
using System.Threading.Tasks;

namespace Palma.GestorDocumental.Repository.Common.SharepointUtils
{
    public class SPOHelper
    {

        public delegate T Mapear<T>(SP.ListItem Items);

        public static SP.Web getWebParentSite(SP.ClientContext clientContext, string consumerSiteLevel)
        {
            SP.Web web = clientContext.Web;
            switch (consumerSiteLevel)
            {
                case obj.SiteSmad:
                    clientContext.Load(web, website => website.Webs, website => website.Title);
                    clientContext.ExecuteQuery();
                    web = web.Webs[1];
                    break;
                case obj.Root:
                    web = clientContext.Site.RootWeb;
                    break;
                default:
                    break;
            }
            return web;
        }

        public static SP.ListItemCollection GetLisItemsParentSite(SP.ClientContext clientContext, string nombrelista, string consumerSiteLevel, string queryCaml)
        {
            SP.Web web = getWebParentSite(clientContext, consumerSiteLevel);

            SP.List lista = web.Lists.GetByTitle(nombrelista);
            SP.CamlQuery query = new SP.CamlQuery();
            query.ViewXml = queryCaml;
            SP.ListItemCollection items = lista.GetItems(query);
            clientContext.Load(items);
            clientContext.ExecuteQuery();
            return items;
        }

        public static List<SP.ListItem> GetLisItemsParentSiteRecursive(SP.ClientContext clientContext, string nombrelista, string consumerSiteLevel, string queryCaml)
        {
            List<SP.ListItem> items = new List<SP.ListItem>();
            SP.Web web = getWebParentSite(clientContext, consumerSiteLevel);
            SP.List lista = web.Lists.GetByTitle(nombrelista);
            SP.ListItemCollectionPosition position = null;
            // Page Size: 100
            int rowLimit = 1000;
            var camlQuery = new SP.CamlQuery();
            camlQuery.ViewXml = queryCaml.Replace("[%ROW_LIMIT%]", rowLimit.ToString());
            do
            {
                SP.ListItemCollection listItems = null;
                camlQuery.ListItemCollectionPosition = position;
                listItems = lista.GetItems(camlQuery);
                clientContext.Load(listItems);
                clientContext.ExecuteQuery();
                position = listItems.ListItemCollectionPosition;
                items.AddRange(listItems.ToList());
            }
            while (position != null);
            return items;
        }

        public static object ObtenerDatosListaMultipleCamposCAMLSinFecha(SP.ClientContext context, object usuario, List<string> diccionario, string v)
        {
            throw new NotImplementedException();
        }

        public static SP.ListItemCollection GetLisItemsParentSiteRoot(SP.ClientContext clientContext, string nombrelista)
        {
            SP.List lista = clientContext.Site.RootWeb.Lists.GetByTitle(nombrelista);
            string queryCAML = @"<View>
            <Query>
            </Query>
            </View>";
            SP.CamlQuery query = new SP.CamlQuery();
            query.ViewXml = queryCAML;
            SP.ListItemCollection items = lista.GetItems(query);
            clientContext.Load(items);
            clientContext.ExecuteQuery();
            return items;
        }

        public static SP.ListItem GetLisItemsbyCaml(SP.ClientContext clientContext, string nombrelista, string queryCaml)
        {
            SP.List lista = clientContext.Web.Lists.GetByTitle(nombrelista);
            SP.CamlQuery query = new SP.CamlQuery();
            query.ViewXml = queryCaml;
            SP.ListItem items = lista.GetItems(query).FirstOrDefault();
            clientContext.Web.Context.Load(items);
            clientContext.Web.Context.ExecuteQuery();
            return items;
        }

        public static SP.ListItem GetListItembyAnyCondition(SP.ClientContext clientContext, string nombrelista, int itemID, string queryCaml)
        {
            SP.List lista = clientContext.Web.Lists.GetByTitle(nombrelista);
            SP.CamlQuery query = new SP.CamlQuery();
            query.ViewXml = queryCaml;
            SP.ListItem items = lista.GetItems(query).FirstOrDefault();
            clientContext.Load(items);
            clientContext.ExecuteQuery();
            return items;
        }

        public static SP.ListItemCollection GetLisItems(SP.ClientContext clientContext, string nombrelista, string queryCAML)
        {
            SP.List lista = clientContext.Web.Lists.GetByTitle(nombrelista);
            SP.CamlQuery query = new SP.CamlQuery();
            query.ViewXml = queryCAML;
            SP.ListItemCollection items = lista.GetItems(query);
            clientContext.Load(items);
            clientContext.ExecuteQuery();
            return items;
        }

        public static SP.ListItemCollection GetLisUsers(SP.ClientContext clientContext, string nombrelista, string queryCAML)
        {
            SP.List lista = clientContext.Web.Lists.GetByTitle(nombrelista);
            SP.CamlQuery query = new SP.CamlQuery();
            query.ViewXml = queryCAML;
            SP.ListItemCollection items = lista.GetItems(query);
            clientContext.Load(items);
            clientContext.ExecuteQuery();
            return items;
        }

        public static SP.ListItemCollection GetListItemsAdvance(string listName, string query, string joins, string projectedFields, string viewFields, string pagingInfo, SP.ClientContext clientContext, uint rowsPerPage = 10)
        {
            SP.ListItemCollection itemCollection = null;
            SP.Web web = clientContext.Web;
            SP.List lista = web.Lists.GetByTitle(listName);
            SP.CamlQuery spQuery = new SP.CamlQuery();
            SP.ListItemCollectionPosition position = null;
            if (!string.IsNullOrEmpty(pagingInfo))
            {
                position = new SP.ListItemCollectionPosition();
                position.PagingInfo = pagingInfo;
                spQuery.ListItemCollectionPosition = position;
            }
            else
                spQuery.ListItemCollectionPosition = position;
            spQuery.ViewXml = query;
            itemCollection = lista.GetItems(spQuery);
            clientContext.Load(itemCollection);
            clientContext.ExecuteQuery();
            return itemCollection;
        }

        public static SP.ListItem GetLisItemsbyTitle(SP.ClientContext clientContext, string nombrelista, string Title)
        {
            SP.List lista = clientContext.Web.Lists.GetByTitle(nombrelista);
            string queryCAML = @"<View>
                <Query>
                    <Where>
                        <Eq>
                            <FieldRef Name='Title'/>
                            <Value Type='Text'>" + Title + @"</Value>
                        </Eq>                
                    </Where>
                    <RowLimit>1</RowLimit>
                </Query>
            </View>";
            SP.CamlQuery query = new SP.CamlQuery();
            query.ViewXml = queryCAML;
            SP.ListItemCollection items = lista.GetItems(query);
            clientContext.Web.Context.Load(items);
            clientContext.Web.Context.ExecuteQuery();
            if (items.Count > 0)
                return items[0];
            else
                return null;
        }

        public static SP.ListItemCollection GetLisItemsbyContainsTitle(SP.ClientContext clientContext, string nombrelista, string Title)
        {
            SP.List lista = clientContext.Web.Lists.GetByTitle(nombrelista);
            string queryCAML = @"<View>
            <Query>
              <Where>
                <Contains>
                  <FieldRef Name='Title'/>
                  <Value Type='Text'>" + Title + @"</Value>
                </Contains>                
              </Where>
            </Query>
          </View>";
            SP.CamlQuery query = new SP.CamlQuery();
            query.ViewXml = queryCAML;
            SP.ListItemCollection items = lista.GetItems(query);
            clientContext.Web.Context.Load(items);
            clientContext.Web.Context.ExecuteQuery();
            return items;
        }

        public static SP.ListItem UploadDocument(SP.ClientContext clientContext, string URLListName, string documentListName, string documentName, byte[] documentStream)
        {
            SP.Web web = clientContext.Web;
            clientContext.Load(web);
            clientContext.ExecuteQuery();
            SP.List documentsList = web.Lists.GetByTitle(documentListName);
            var fileCreationInformation = new SP.FileCreationInformation();
            fileCreationInformation.Content = documentStream;
            fileCreationInformation.Overwrite = true;
            fileCreationInformation.Url = web.Url + "/" + URLListName + "/" + documentName;
            SP.File uploadFile = documentsList.RootFolder.Files.Add(
                fileCreationInformation);
            uploadFile.ListItemAllFields.Update();
            clientContext.ExecuteQuery();
            return uploadFile.ListItemAllFields;
        }

        public static SP.ListItem UploadDocument2(SP.ClientContext clientContext, string documentListName, string documentName, Stream documentStream)
        {
            SP.List documentsList = clientContext.Web.Lists.GetByTitle(documentListName);
            clientContext.Load(documentsList);
            var fileCreationInformation = new SP.FileCreationInformation();
            fileCreationInformation.ContentStream = documentStream;
            fileCreationInformation.Overwrite = true;
            string[] strArray = documentName.Split('\\');
            fileCreationInformation.Url = strArray[strArray.Length - 1];
            SP.File uploadFile = documentsList.RootFolder.Files.Add(fileCreationInformation);
            clientContext.ExecuteQuery();
            return uploadFile.ListItemAllFields;
        }

        public static SP.ListItem UploadDocument3(SP.ClientContext clientContext, string consumerSiteLevel, string documentListName, string documentName, Stream documentStream, Dictionary<string, object> CampoValor)
        {
            SP.Web web = getWebParentSite(clientContext, consumerSiteLevel);
            SP.List documentsList = web.Lists.GetByTitle(documentListName);
            clientContext.Load(documentsList);
            var fileCreationInformation = new SP.FileCreationInformation();
            fileCreationInformation.ContentStream = documentStream;
            fileCreationInformation.Overwrite = true;
            string[] strArray = documentName.Split('\\');
            fileCreationInformation.Url = strArray[strArray.Length - 1];
            SP.File uploadFile = documentsList.RootFolder.Files.Add(fileCreationInformation);
            clientContext.ExecuteQuery();
            SP.ListItem currentItem = uploadFile.ListItemAllFields;
            foreach (KeyValuePair<string, object> kvp in CampoValor)
            {
                currentItem[kvp.Key] = kvp.Value;
            }
            currentItem.Update();
            clientContext.Load(currentItem);
            clientContext.ExecuteQuery();
            return currentItem;
        }

        public static SP.ListItem UploadDocument4(SP.ClientContext clientContext, string consumerSiteLevel, string documentListName, string documentName, byte[] documentStream, Dictionary<string, object> CampoValor)
        {
            SP.Web web = getWebParentSite(clientContext, consumerSiteLevel);
            SP.List documentsList = web.Lists.GetByTitle(documentListName);
            clientContext.Load(documentsList);
            var fileCreationInformation = new SP.FileCreationInformation();
            fileCreationInformation.Content = documentStream;
            fileCreationInformation.Overwrite = true;
            string[] strArray = documentName.Split('\\');
            fileCreationInformation.Url = strArray[strArray.Length - 1];
            SP.File uploadFile = documentsList.RootFolder.Files.Add(fileCreationInformation);
            clientContext.ExecuteQuery();
            SP.ListItem currentItem = uploadFile.ListItemAllFields;
            foreach (KeyValuePair<string, object> kvp in CampoValor)
            {
                currentItem[kvp.Key] = kvp.Value;
            }
            currentItem.Update();
            clientContext.Load(currentItem);
            clientContext.ExecuteQuery();
            return currentItem;
        }

        public static bool HasElements(SP.ClientContext clientContext, string nombrelista, string consumerSiteLevel, string queryCaml)
        {
            SP.Web web = getWebParentSite(clientContext, consumerSiteLevel);
            SP.List lista = web.Lists.GetByTitle(nombrelista);
            SP.CamlQuery query = new SP.CamlQuery();
            query.ViewXml = queryCaml;
            SP.ListItemCollection items = lista.GetItems(query);
            clientContext.Load(items);
            clientContext.ExecuteQuery();
            return items.Count == 0;
        }

        public static SP.User ObtenerUsuarioxID(SP.ClientContext context, int IdUser)
        {
            try
            {
                SP.User user = context.Web.GetUserById(IdUser);
                context.Load(user, u => u.Id, u => u.Title, u => u.Email, u => u.LoginName);
                context.ExecuteQuery();
                return user;
            }
            catch (Exception)
            {
                return null;
            }
        }

        public static SP.User ObtenerUsuarioxEmail(SP.ClientContext clientContext, string Email)
        {
            try
            {
                SP.ClientResult<PrincipalInfo> persons = SP.Utilities.Utility.ResolvePrincipal(clientContext, clientContext.Web, Email, PrincipalType.User, PrincipalSource.All, null, true);
                clientContext.ExecuteQuery();
                PrincipalInfo person = persons.Value;
                SP.Web web = clientContext.Web;
                SP.User user = web.EnsureUser(person.LoginName);
                clientContext.ExecuteQuery();
                user = clientContext.Site.RootWeb.SiteUsers.GetByEmail(Email);
                clientContext.Load(user, u => u.Id, u => u.Title, u => u.Email, u => u.LoginName);
                clientContext.ExecuteQuery();
                return user;
            }
            catch (Exception ex)
            {
                return null;
            }
        }

        public static bool esSuperAdmin(SP.ClientContext clientContext, string LoginName, string Email)
        {
            bool isSuperAdmin = false;
            try
            {

                SP.Web web = clientContext.Site.RootWeb;
                SP.User user = web.EnsureUser(LoginName);
                clientContext.ExecuteQuery();
                user = clientContext.Site.RootWeb.SiteUsers.GetByEmail(Email);
                clientContext.Load(user, u => u.IsSiteAdmin);
                clientContext.ExecuteQuery();
                isSuperAdmin = user.IsSiteAdmin;
            }
            catch (Exception e)
            {
                isSuperAdmin = false;
                throw e;
            }
            return isSuperAdmin;
        }

        public static SP.User ObtenerUsuarioxLoginName(SP.ClientContext context, string loginname)
        {
            try
            {
                SP.User user = context.Web.SiteUsers.GetByLoginName(loginname);
                context.Load(user);
                context.ExecuteQuery();
                return user;
            }
            catch (Exception)
            {
                return null;
            }
        }

        public static SP.User ObtenerUsuarioLogin(SP.ClientContext clientContext)
        {
            try
            {
                SP.User spUser = clientContext.Web.CurrentUser;
                clientContext.Load(spUser, user => user.Title, user => user.Email, user => user.Id, user => user.LoginName, user => user.Groups);
                clientContext.ExecuteQuery();
                return spUser;
            }
            catch (Exception)
            {
                return null;
            }
        }

        public static bool ChekUserInGroup(int IDUsario, string NombreGrupo, SP.ClientContext clientContextWeb)
        {
            bool retorno = false;

            try
            {
                SP.GroupCollection siteGroups = clientContextWeb.Web.SiteGroups;
                SP.User spUser = clientContextWeb.Web.GetUserById(IDUsario);
                SP.Group membersGroup = siteGroups.GetByName(NombreGrupo);
                clientContextWeb.Load(membersGroup);
                clientContextWeb.Load(spUser, user => user.Groups);
                clientContextWeb.ExecuteQuery();
                foreach (var grupo in spUser.Groups)
                {
                    if (grupo.Id == membersGroup.Id)
                    {
                        retorno = true;
                        break;
                    }
                }
            }
            catch (Exception)
            {
            }
            return retorno;
        }

        public static SP.CamlQuery CreateJoinQuery(string joinListTitle, string joinFieldName, string[] viewdFields, string[] projectedFields)
        {
            var qry = new SP.CamlQuery();
            qry.ViewXml = @"<View>
               <ViewFields>";
            foreach (var f in viewdFields)
            {
                qry.ViewXml += string.Format("<FieldRef Name='{0}' />", f);
            }
            foreach (var f in projectedFields)
            {
                qry.ViewXml += string.Format("<FieldRef Name='{0}{1}' />", joinListTitle, f);
            }
            qry.ViewXml += @"</ViewFields>
               <ProjectedFields>";
            foreach (var f in projectedFields)
            {
                qry.ViewXml += string.Format("<Field Name='{0}{1}' Type='Lookup' List='{0}' ShowField='{1}' />", joinListTitle, f);
            }
            qry.ViewXml += string.Format(@"</ProjectedFields>
               <Joins>
                   <Join Type='LEFT' ListAlias='{0}'>
                       <Eq>
                           <FieldRef Name='{1}' RefType='ID' />
                           <FieldRef List='{0}' Name='ID' />
                       </Eq>
                   </Join>
               </Joins>
           </View>", joinListTitle, joinFieldName);
            return qry;
        }

        public static int AddItem(SP.ClientContext clientContext, string nombrelista, string consumerSiteLevel, Dictionary<string, object> CampoValor)
        {
            var returnId = 0;
            try
            {
                SP.Web web = getWebParentSite(clientContext, consumerSiteLevel);
                var myList = web.Lists.GetByTitle(nombrelista);
                SP.ListItem newItem = myList.AddItem(new SP.ListItemCreationInformation());
                foreach (KeyValuePair<string, object> kvp in CampoValor)
                {
                    newItem[kvp.Key] = kvp.Value;
                }
                newItem.Update();
                clientContext.ExecuteQuery();
                returnId = newItem.Id;
            }
            catch (Exception ex)
            {
                throw ex;
            }
            return returnId;
        }

        public static int AddList(SP.ClientContext clientContext, string listName, string consumerSiteLevel, List<Dictionary<string, object>> dictionaryList)
        {
            var returnId = 0;
            SP.Web web = getWebParentSite(clientContext, consumerSiteLevel);
            SP.List list = web.Lists.GetByTitle(listName);
            SP.ListItem listItem = null;
            try
            {
                foreach (var dictionary in dictionaryList)
                {
                    listItem = list.AddItem(new SP.ListItemCreationInformation());
                    foreach (KeyValuePair<string, object> kvp in dictionary)
                    {
                        listItem[kvp.Key] = kvp.Value;
                    }
                    listItem.Update();
                    clientContext.Load(listItem);
                }
                clientContext.ExecuteQuery();
                returnId = listItem.Id;
            }
            catch (Exception ex)
            {
                throw ex;
            }
            return returnId;
        }

        public static int DeleteWithQuery(SP.ClientContext context, string listName, string camlQuery)
        {
            var returnId = 0;
            SP.Web web = getWebParentSite(context, string.Empty);
            try
            {
                SP.List list = web.Lists.GetByTitle(listName);
                var query = new SP.CamlQuery { ViewXml = camlQuery };
                SP.ListItemCollection itemsCollection = list.GetItems(query);
                web.Context.Load(itemsCollection);
                web.Context.ExecuteQuery();
                var listToDelete = new List<SP.ListItem>();
                foreach (var item in itemsCollection)
                {
                    listToDelete.Add(item);
                }
                foreach (var item in listToDelete)
                {
                    item.DeleteObject();
                }
                web.Context.ExecuteQuery();
            }
            catch (Exception ex)
            {
                throw ex;
            }
            return returnId;
        }

        public static int UpdateItem(SP.ClientContext clientContext, string nombrelista, string consumerSiteLevel, int itemID, Dictionary<string, object> CampoValor)
        {
            var returnId = 0;
            try
            {
                SP.Web web = getWebParentSite(clientContext, consumerSiteLevel);
                var myList = web.Lists.GetByTitle(nombrelista);
                SP.ListItem currentItem = myList.GetItemById(itemID);
                foreach (KeyValuePair<string, object> kvp in CampoValor)
                {
                    currentItem[kvp.Key] = kvp.Value;
                }
                currentItem.Update();
                clientContext.ExecuteQuery();
                returnId = itemID;
            }
            catch (Exception ex)
            {
                throw ex;
            }

            return returnId;
        }

        public static int UpdateItemWithQueryCaml(SP.ClientContext clientContext, string nombrelista, string consumerSiteLevel, string query, Dictionary<string, object> CampoValor)
        {
            var returnId = 0;
            try
            {
                SP.Web web = getWebParentSite(clientContext, consumerSiteLevel);
                var myList = web.Lists.GetByTitle(nombrelista);
                var querycaml = new SP.CamlQuery { ViewXml = query };
                SP.ListItemCollection itemsCollection = myList.GetItems(querycaml);
                web.Context.Load(itemsCollection);
                web.Context.ExecuteQuery();
                var currentItem = itemsCollection.FirstOrDefault();
                if (currentItem != null)
                {
                    foreach (KeyValuePair<string, object> kvp in CampoValor)
                    {
                        currentItem[kvp.Key] = kvp.Value;
                    }
                    currentItem.Update();
                    web.Context.ExecuteQuery();
                    returnId = (int)currentItem["ID"];
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }

            return returnId;
        }

        public static int UpdateListWithQueryCaml(SP.ClientContext clientContext, string nombrelista, string consumerSiteLevel, string query, Dictionary<string, object> CampoValor)
        {
            var returnId = 0;
            try
            {
                SP.Web web = getWebParentSite(clientContext, consumerSiteLevel);
                var myList = web.Lists.GetByTitle(nombrelista);
                var querycaml = new SP.CamlQuery { ViewXml = query };
                SP.ListItemCollection itemsCollection = myList.GetItems(querycaml);
                web.Context.Load(itemsCollection);
                web.Context.ExecuteQuery();
                foreach (var currentItem in itemsCollection)
                {
                    foreach (KeyValuePair<string, object> kvp in CampoValor)
                    {
                        currentItem[kvp.Key] = kvp.Value;
                    }
                    currentItem.Update();
                    web.Context.ExecuteQuery();
                    returnId = (int)currentItem["ID"];
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }

            return returnId;
        }

        public static int DeleteItembyID(SP.ClientContext clientContext, string nombreLista, string consumerSiteLevel, string queryCAML, int IDElemento)
        {
            var returnId = 0;
            try
            {
                SP.ListItem listItem;
                SP.Web web = getWebParentSite(clientContext, consumerSiteLevel);
                SP.List documentsList = web.Lists.GetByTitle(nombreLista);
                if (!string.IsNullOrEmpty(queryCAML))
                {
                    SP.CamlQuery query = new SP.CamlQuery();
                    query.ViewXml = queryCAML;
                    listItem = documentsList.GetItems(query).FirstOrDefault();
                }
                else
                {
                    listItem = documentsList.GetItemById(IDElemento);
                }
                listItem.DeleteObject();
                clientContext.ExecuteQuery();
                returnId = IDElemento;
            }
            catch (Exception ex)
            {
                throw ex;
            }
            return returnId;
        }

        public static List<SP.User> ObtenerUsuariosDeGrupo(string NombreGrupo, SP.ClientContext clientContextWeb)
        {
            List<SP.User> lsUsers = new List<SP.User>();
            SP.GroupCollection siteGroups = clientContextWeb.Web.SiteGroups;
            SP.Group membersGroup = siteGroups.GetByName(NombreGrupo);
            clientContextWeb.Load(membersGroup, x => x.Users);
            clientContextWeb.ExecuteQuery();
            foreach (var member in membersGroup.Users)
            {
                lsUsers.Add(member);
            }
            return lsUsers;
        }

        public static void DeleteAllLisItems(SP.ClientContext clientContext, string nombrelista)
        {
            SP.List lista = clientContext.Web.Lists.GetByTitle(nombrelista);
            SP.ListItemCollectionPosition licp = null;
            clientContext.Load(lista);
            clientContext.ExecuteQuery();
            while (true)
            {
                SP.CamlQuery query = new SP.CamlQuery();
                query.ViewXml = @"<View><Query><Where><IsNotNull><FieldRef Name='ID' /></IsNotNull></Where></Query>
                                <RowLimit>250</RowLimit><ViewFields><FieldRef Name='Id'/></ViewFields></View>";
                query.ListItemCollectionPosition = licp;
                SP.ListItemCollection items = lista.GetItems(query);
                clientContext.Load(items);
                clientContext.ExecuteQuery();
                int cantidad = items.Count;
                Console.WriteLine("Elementos: " + cantidad);
                licp = items.ListItemCollectionPosition;
                foreach (var item in items.ToList())
                {
                    item.DeleteObject();
                }
                clientContext.ExecuteQuery();
                if (licp == null)
                    break;
            }
        }

        public static int EnviarEmailSharePoint(SP.ClientContext clientContext, List<string> lstDestinatarios, string _Asunto, string _Mensaje, List<string> lstCorreosCopia = null)
        {
            int IdGenerado = 0;
            try
            {
                SP.Utilities.EmailProperties emailprop = new SP.Utilities.EmailProperties();
                emailprop.To = lstDestinatarios;
                emailprop.Subject = _Asunto;
                emailprop.Body = _Mensaje;
                if (lstCorreosCopia != null && lstCorreosCopia.Count > 0)
                    emailprop.CC = lstCorreosCopia;
                SP.Utilities.Utility.SendEmail(clientContext, emailprop);
                clientContext.ExecuteQuery();
                IdGenerado = 1;
            }
            catch (Exception ex)
            {
                throw ex;
            }
            return IdGenerado;
        }

        public static bool RomperHerenciaxItem(SP.ClientContext clientContext, string nombreLista, int ItemId)
        {
            bool retorno = false;
            try
            {
                var myList = clientContext.Web.Lists.GetByTitle(nombreLista);
                SP.ListItem currentItem = myList.GetItemById(ItemId);
                clientContext.Load(currentItem);
                clientContext.ExecuteQuery();
                currentItem.BreakRoleInheritance(true, false);
                currentItem.Update();
                retorno = true;
            }
            catch (Exception ex)
            {
                throw ex;
            }
            return retorno;
        }

        public static void EliminarTodosLosPermisos(SP.ClientContext clientContext, string nombreLista, int ItemId)
        {
            try
            {
                var myList = clientContext.Web.Lists.GetByTitle(nombreLista);
                SP.ListItem currentItem = myList.GetItemById(ItemId);
                clientContext.Load(currentItem, Item => Item.RoleAssignments);
                clientContext.ExecuteQuery();
                for (int i = currentItem.RoleAssignments.Count - 1; i >= 0; --i)
                {
                    currentItem.RoleAssignments[i].DeleteObject();
                    clientContext.ExecuteQuery();
                }
            }
            catch (Exception ex)
            {

                throw ex;
            }
        }

        public static void AgregarUsuarioPermisosxElemento(SP.ClientContext clientContext, string nombreLista, int ItemId, string loginname, string levelPermission)
        {
            try
            {
                var myList = clientContext.Web.Lists.GetByTitle(nombreLista);
                SP.ListItem currentItem = myList.GetItemById(ItemId);
                clientContext.Load(currentItem, Item => Item.RoleAssignments);
                clientContext.Load(clientContext.Web.RoleDefinitions);
                clientContext.ExecuteQuery();
                var roletypes = clientContext.Web.RoleDefinitions.GetByName(levelPermission);
                SP.RoleDefinitionBindingCollection rdb = new SP.RoleDefinitionBindingCollection(clientContext);
                rdb.Add(roletypes);
                SP.Principal usr = clientContext.Web.EnsureUser(loginname);
                currentItem.RoleAssignments.Add(usr, rdb);
                clientContext.ExecuteQuery();
            }
            catch (Exception ex)
            {

                throw ex;
            }
        }

        public static void AgregarListaUsuarioPermisosxElemento(SP.ClientContext clientContext, string nombreLista, int ItemId, List<string> lstLoginname, string levelPermission)
        {
            try
            {
                var myList = clientContext.Web.Lists.GetByTitle(nombreLista);
                SP.ListItem currentItem = myList.GetItemById(ItemId);
                clientContext.Load(currentItem, Item => Item.RoleAssignments);
                clientContext.Load(clientContext.Web.RoleDefinitions);
                clientContext.ExecuteQuery();
                SP.RoleDefinition rd = clientContext.Web.RoleDefinitions.GetByName(levelPermission);
                SP.RoleDefinitionBindingCollection rdb = new SP.RoleDefinitionBindingCollection(clientContext);
                rdb.Add(rd);
                foreach (var loginName in lstLoginname)
                {
                    SP.Principal usr = clientContext.Web.EnsureUser(loginName);
                    currentItem.RoleAssignments.Add(usr, rdb);
                }
                clientContext.ExecuteQuery();
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public static void AgregarGrupoPermisosxElemento(SP.ClientContext clientContext, string nombreLista, int ItemId, string groupName, string levelPermission)
        {
            try
            {
                var myList = clientContext.Web.Lists.GetByTitle(nombreLista);
                SP.ListItem currentItem = myList.GetItemById(ItemId);
                clientContext.Load(currentItem, Item => Item.RoleAssignments);
                clientContext.Load(clientContext.Web.RoleDefinitions);
                clientContext.ExecuteQuery();
                SP.RoleDefinitionCollection roleDefs = clientContext.Web.RoleDefinitions;
                var query = clientContext.LoadQuery(roleDefs.Where(p => p.Name == levelPermission));
                clientContext.ExecuteQuery();
                SP.RoleDefinition roledefObj = query.FirstOrDefault();
                SP.Principal group = ObtenerGrupoPorNombre(clientContext, groupName);
                if (group != null)
                {
                    SP.RoleDefinitionBindingCollection collRoleDefinitionBinding = new SP.RoleDefinitionBindingCollection(clientContext);
                    collRoleDefinitionBinding.Add(roledefObj);
                    var roleAssignments = currentItem.RoleAssignments;
                    roleAssignments.Add(group, collRoleDefinitionBinding);
                    clientContext.ExecuteQuery();
                }

            }
            catch (Exception ex)
            {

                throw ex;
            }
        }

        public static SP.Group ObtenerGrupoPorNombre(SP.ClientContext clientContextWeb, string NombreGrupo)
        {

            SP.GroupCollection siteGroups = clientContextWeb.Web.SiteGroups;
            SP.Group group = siteGroups.GetByName(NombreGrupo);
            clientContextWeb.Load(group);
            clientContextWeb.ExecuteQuery();
            return group;

        }

        public static SP.Group ObtenerGrupoPorId(SP.ClientContext clientContextWeb, int Id)
        {

            SP.GroupCollection siteGroups = clientContextWeb.Web.SiteGroups;
            SP.Group group = siteGroups.GetById(Id);
            clientContextWeb.Load(group);
            clientContextWeb.ExecuteQuery();
            return group;

        }

        public static Stream ObtenerArchivo(SP.ClientContext clientContext, string pe_strUrlFile)
        {
            Stream vr_objStream = null;
            try
            {
                SP.Web web = clientContext.Web;
                SP.File objSPFile = web.GetFileByServerRelativeUrl(pe_strUrlFile);
                clientContext.Load(objSPFile);
                if (objSPFile != null)
                {
                    SP.ClientResult<Stream> streamResult = objSPFile.OpenBinaryStream();
                    clientContext.ExecuteQuery();
                    vr_objStream = streamResult.Value;
                }

            }
            catch (Exception ex)
            {
                throw ex;
            }
            return vr_objStream;
        }

        public static SP.ClientResult<ResultTableCollection> GetItemsSearch(SP.ClientContext clientContext, string query, string[] properties)
        {
            KeywordQuery keywordQuery = new KeywordQuery(clientContext);
            keywordQuery.QueryText = query;
            foreach (string property in properties)
            {
                keywordQuery.SelectProperties.Add(property);
            }
            keywordQuery.TrimDuplicates = false;
            keywordQuery.EnableQueryRules = true;
            keywordQuery.RowLimit = 500;
            SearchExecutor searchExecutor = new SearchExecutor(clientContext);
            SP.ClientResult<ResultTableCollection> results = searchExecutor.ExecuteQuery(keywordQuery);
            clientContext.ExecuteQuery();
            return results;
        }

    }
}
