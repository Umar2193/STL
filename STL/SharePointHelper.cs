namespace STL {
    public class SharePointHelper {

        private readonly string Username;
        private readonly string Password;
        private readonly string SiteUrl;
        private readonly string BasePath;

        public SharePointHelper(string url, string username, string password, string basePath) {
            SiteUrl = url;
            Username = username;
            Password = password;
            BasePath = basePath;
        }

        public void SaveFile(string filePath, string sharePointPath) {

            System.IO.File.Copy(filePath, sharePointPath);

            //try {
            //    var securePassword = new SecureString();
            //    foreach(char c in Password) { securePassword.AppendChar(c); }

            //    Uri site = new Uri(SiteUrl);
            //    using(var authenticationManager = new AuthenticationManager())
            //    using(var CContext = authenticationManager.GetContext(site, Username, securePassword)) {
            //        Web web = CContext.Web;
            //        FileCreationInformation newFile = new FileCreationInformation();
            //        byte[] FileContent = System.IO.File.ReadAllBytes(filePath);
            //        newFile.ContentStream = new MemoryStream(FileContent);
            //        newFile.Url = Path.GetFileName(filePath);
            //        List DocumentLibrary = web.Lists.GetByTitle(sharePointPath);
            //        Folder Clientfolder = null;
            //        if(sharePointPath == "") {
            //            Clientfolder = DocumentLibrary.RootFolder;
            //        } else {
            //            Clientfolder = DocumentLibrary.RootFolder.Folders.Add(sharePointPath);
            //            Clientfolder.Update();
            //        }
            //        Microsoft.SharePoint.Client.File uploadFile = Clientfolder.Files.Add(newFile);
            //        CContext.Load(DocumentLibrary);
            //        CContext.Load(uploadFile);
            //        CContext.ExecuteQuery();
            //    }
            //} catch(Exception exp) {

            //} 
        }

        public void moveFile(string oldPath, string newPath) {
            System.IO.File.Move(oldPath, newPath);
        }

        public void CreateFolder(string name) {
            if(!Directory.Exists($"{BasePath}\\{name}")) {
                Directory.CreateDirectory($"{BasePath}\\{name}");
            }
        }

        public void MoveFolder(string oldName, string newName) {
            if(Directory.Exists($"{BasePath}\\{oldName}")) {
                Directory.Move($"{BasePath}\\{oldName}", $"{BasePath}\\{newName}");
            }
        }

        public bool CheckIfFolderExists(string name) {
            return Directory.Exists($"{BasePath}\\{name}");
        }

        public IEnumerable<string> ReadFolders(string sharePointPath) {
            try {

                return Directory.GetDirectories(BasePath).Select(c => c.Replace($"{BasePath}\\", ""));
                //var securePassword = new SecureString();
                //foreach(char c in Password) { securePassword.AppendChar(c); }

                //Uri site = new Uri(SiteUrl + sharePointPath);
                //using(var authenticationManager = new AuthenticationManager())
                //using(var CContext = authenticationManager.GetContext(site, Username, securePassword)) {
                //    var list = new List<string>();
                //    Web web = CContext.Web;

                //    // Retrieve all lists from the server, and put the return value in another
                //    // collection instead of the web.Lists.
                //    IEnumerable<List> result = CContext.LoadQuery(
                //      web.Lists.Include(
                //          // For each list, retrieve Title and Id.
                //          list => list.Title,
                //          list => list.Id
                //      )
                //    );

                //    // Execute query.
                //    CContext.ExecuteQuery();

                //    // Enumerate the result.
                //    foreach(List tmp in result) {
                //        list.Add(tmp.Title);
                //    }
                //    return list;
                // }
            } catch(Exception exp) {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine(exp.Message + Environment.NewLine + exp.StackTrace);
            } finally {
                Console.ReadLine();
            }
            return null;
        }
    }
}
