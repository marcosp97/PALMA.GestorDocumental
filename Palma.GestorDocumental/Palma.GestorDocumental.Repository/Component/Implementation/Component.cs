using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.Extensions.Configuration;
using Microsoft.SharePoint.Client;
using Palma.GestorDocumental.Repository.Common.FileUtils;
using Palma.GestorDocumental.Repository.Common.SharepointUtils;
using Palma.GestorDocumental.Repository.Component.Interface;
using Palma.GestorDocumental.Repository.Connection;
using Palma.GestorDocumental.Repository.Entity;
using obj = Palma.GestorDocumental.Repository.Common.Configuration.Manager.ObjectSharepoint.Site;
using util = Palma.GestorDocumental.Repository.Common.Helpers.ObjectHelper;

namespace Palma.GestorDocumental.Repository.Component.Implementation
{
    public class Component : IComponent
    {
        private readonly IConfiguration _configuration;
        public Component(
            IConfiguration configuration
            )
        {
            this._configuration = configuration;
        }
        public void Run()
        {
            string user = this._configuration["Credentials:user"];
            string pass = this._configuration["Credentials:password"];
            string siteUrl = this._configuration["Credentials:urlSHP"];
            string libraryElaboracion = this._configuration["Credentials:libraryEnElaboracion"];
            string libraryObsoletos = this._configuration["Credentials:libraryObsoletos"];
            string libraryVigentes = this._configuration["Credentials:libraryVigentes"];
            string oldLink = this._configuration["Credentials:oldLink"];
            string newLink = this._configuration["Credentials:newLink"];
            SecureString password = new SecureString();
            pass.ToList().ForEach(password.AppendChar);

            Uri site = new Uri(siteUrl);

            using (var authenticationManager = new AuthenticationManager())
            using (var context = authenticationManager.GetContext(site, user, password))
            {
                string camlQuery = @"<View Scope='RecursiveAll'>
                                        <Query>
                                            <Where>
                                                <Or>
                                                    <IsNull>
                                                    <FieldRef Name='LinkCambiado' />
                                                    </IsNull>
                                                    <Eq>
                                                    <FieldRef Name='LinkCambiado' />
                                                    <Value Type='Boolean'>false</Value>
                                                    </Eq>
                                                </Or>
                                            </Where>
                                            <OrderBy Override='TRUE'><FieldRef Name='ID' Ascending='False'/></OrderBy>
                                        </Query>
                                        <RowLimit Paged='TRUE'>[%ROW_LIMIT%]</RowLimit></View>";
                List<FileBE> elementsElaboracion = SPOHelper.GetLisItemsParentSiteRecursive(context, libraryElaboracion, obj.Root, camlQuery).Select(x => new FileBE() { item = x, nombreLista = libraryElaboracion }).ToList();
                List<FileBE> elementsObsoletos = SPOHelper.GetLisItemsParentSiteRecursive(context, libraryObsoletos, obj.Root, camlQuery).Select(x => new FileBE() { item = x, nombreLista = libraryObsoletos }).ToList();
                List<FileBE> elementsVigentes = SPOHelper.GetLisItemsParentSiteRecursive(context, libraryVigentes, obj.Root, camlQuery).Select(x => new FileBE() { item = x, nombreLista = libraryVigentes }).ToList();

                List<FileBE> elements = new List<FileBE>();
                elements.AddRange(elementsElaboracion);
                elements.AddRange(elementsObsoletos);
                elements.AddRange(elementsVigentes);

                //elements.ForEach(element =>
                //{
                //    var fileRef = util.toString(element["FileRef"]);
                //    try
                //    {

                //        Stream file = SPOHelper.ObtenerArchivo(context, fileRef);
                //        Stream newfile = FileHelper.ReplaceHiperLink(file, oldLink, newLink);
                //    }
                //    catch (Exception ex)
                //    {
                //        Console.WriteLine(@$"Ocurrio un error en el archivo {fileRef}: {ex.Message}");
                //    }

                //});

                using var semaforo = new SemaphoreSlim(1);
                var tasks = elements.Select(element => Task.Run(async () =>
                {
                    var fileRef = util.toString(element.item["FileRef"]);
                    var fileName = Path.GetFileName(fileRef);
                    var ID = util.toInt(element.item["ID"]);
                    Dictionary<string, object> dict = new Dictionary<string, object>();
                    dict["LinkCambiado"] = true;
                    await semaforo.WaitAsync();
                    try
                    {
                        Stream file = SPOHelper.ObtenerArchivo(context, fileRef);
                        string newfile = FileHelper.ReplaceHiperLink(file, oldLink, newLink);
                        var bytes = Convert.FromBase64String(newfile);
                        SPOHelper.UploadDocument4(context, obj.Root ,element.nombreLista, fileName, bytes, dict);
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(@$"Ocurrio un error en el archivo {fileRef}: {ex.Message}");
                    }
                    finally
                    {
                        semaforo.Release();
                    }
                })).ToArray();

                Task.WaitAll(tasks);
            }
        }
    }
}
