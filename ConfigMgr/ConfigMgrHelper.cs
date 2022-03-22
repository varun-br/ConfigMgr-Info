using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Management;

namespace ConfigMgr
{
    public class ConfigMgrHelper
    {
        public string Server { get; set; }
        public string UserName { get; set; }
        public string Password { get; set; }
        public string SiteName { get; set; }
        public ManagementScope Scope { get; protected set; }

        public const string Name = "Name";
        public const string ContainerNodeId = "ContainerNodeID";
        public const string ParentContainerNodeId = "ParentContainerNodeID";
        public const string InstanceKey = "InstanceKey";
        public const string LocalizedDisplayName = "LocalizedDisplayName";
        public const string ModelName = "ModelName";
        public const string CiId = "CI_ID";

        public ConfigMgrHelper(string server, string siteName, string userName, string password)
        {
            Server = server;
            UserName = userName;
            Password = password;
            SiteName = siteName;
        }

        public bool Connect()
        {
            try
            {
                if (Scope?.IsConnected ?? false) return true;

                var sw = Stopwatch.StartNew();
                WriteLine("Connecting to ConfigMgr...");

                var path = new ManagementPath($@"\\{Server}\root\Sms\site_{SiteName}");
                var connectOptions = new ConnectionOptions()
                {
                    Username = UserName,
                    Password = Password,
                    Authentication = AuthenticationLevel.PacketPrivacy,
                    Impersonation = ImpersonationLevel.Impersonate
                };
                Scope = new ManagementScope(path, connectOptions);
                Scope.Connect();
                if (Scope.IsConnected)
                {
                    sw.Stop();
                    WriteLine($"Connected to ConfigMgr : {sw.Elapsed}");
                    var version = GetVersion();
                    WriteLine($"ConfigMgr Version : {version}");
                    return true;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
                throw;
            }
            return false;
        }

        public ManagementObjectCollection ExecuteQuery(string query)
        {
            try
            {
                if (Connect())
                {
                    using (var searcher = new ManagementObjectSearcher(Scope, new ObjectQuery(query)))
                    {
                        return searcher.Get();
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
            }
            return null;
        }

        public Version GetVersion()
        {
            var version = "0.0";
            ManagementObjectCollection objPackages = ExecuteQuery("SELECT Version FROM SMS_Site");
            var i = 1;
            foreach (ManagementObject objPackage in objPackages)
            {
                version = objPackage["Version"].ToString().Trim();
                WriteLine($"{i++} - {version}");
            }

            return new Version(version);
        }

        public void BuildApplicationTree(string objectType)
        {
            Connect();
            var sw = Stopwatch.StartNew();
            //Adding Folder Groups.
            WriteLine("Getting Folders...");
            var nodes = new List<Node>();
            var query = $"SELECT {Name},{ContainerNodeId},{ParentContainerNodeId} FROM SMS_ObjectContainerNode WHERE ObjectType = {objectType}";
            var objectCollection = ExecuteQuery(query);

            foreach (var obj in objectCollection)
            {
                var nodeName = obj[Name].ToString();
                var containerId = obj[ContainerNodeId].ToString();
                var parent = obj[ParentContainerNodeId].ToString();
                nodes.Add(new Node(containerId, nodeName, parent));
            }
            
            sw.Stop();
            WriteLine($"Folders loaded: {sw.Elapsed}");

            sw = Stopwatch.StartNew();
            WriteLine("Getting Containers...");
            var containerIds = string.Join(",", nodes.Select(x => x.Id));
            query = $"SELECT {InstanceKey}, {ContainerNodeId} FROM SMS_ObjectContainerItem WHERE {ContainerNodeId} IN ({containerIds})";
            objectCollection = ExecuteQuery(query);
            var containerItems = new Dictionary<string, string>(StringComparer.InvariantCultureIgnoreCase);
            foreach (var obj in objectCollection)
            {
                var instanceKey = obj[InstanceKey].ToString();
                var containerNodeId = obj[ContainerNodeId].ToString();
                if (!containerItems.ContainsKey(instanceKey))
                    containerItems.Add(instanceKey, containerNodeId);
            }

            sw.Stop();
            WriteLine($"Containers loaded: {sw.Elapsed}");

            sw = Stopwatch.StartNew();
            WriteLine("Getting Applications...");

            query = $"SELECT {ModelName},{LocalizedDisplayName},{CiId} FROM SMS_Application WHERE IsLatest='True'";
            objectCollection = ExecuteQuery(query);
            foreach (var obj in objectCollection)
            {
                var modelName = obj[ModelName].ToString();
                string containerId;
                if (!containerItems.TryGetValue(modelName, out containerId))
                    containerId = "0";
                var displayName = obj[LocalizedDisplayName].ToString();
                var id = obj[CiId].ToString();
                nodes.Add(new Node(id, displayName, containerId, modelName));
            }

            sw.Stop();
            WriteLine($"Applications loaded: {sw.Elapsed}");

            WriteLine($"Folders: {nodes.Count(i => !i.IsApp)}");
            WriteLine($"Applications: {nodes.Count(i => i.IsApp)}");
            WriteLine($"Total: {nodes.Count}");
        }

        public List<Node> BuildApplicationTreeRecursive(int objectType, string parentID, bool addObjects, bool IsDependencyRequired)
        {
            var nodes = new List<Node>();
            var query = $"SELECT Name,ContainerNodeID FROM SMS_ObjectContainerNode WHERE ObjectType = {objectType.ToString()} AND ParentContainerNodeID = {parentID.ToString()}";
            var objectCollection = ExecuteQuery(query);

            foreach (var obj in objectCollection)
            {
                var nodeName = obj[Name].ToString();
                var containerId = obj[ContainerNodeId].ToString();
                nodes.Add(new Node(containerId, nodeName, parentID));
                BuildApplicationTreeRecursive(objectType, containerId, addObjects, IsDependencyRequired);
            }

            query = parentID == "0" ? "SELECT ModelName,LocalizedDisplayName,CI_ID FROM SMS_Application WHERE IsLatest='True' AND ModelName NOT IN (SELECT InstanceKey FROM SMS_ObjectContainerItem)"
                : $"SELECT ModelName,LocalizedDisplayName,CI_ID FROM SMS_Application WHERE IsLatest='True' AND ModelName IN (SELECT InstanceKey FROM SMS_ObjectContainerItem WHERE ContainerNodeID = {parentID} )";

            objectCollection = ExecuteQuery(query);
            foreach (var obj in objectCollection)
            {
                var modelName = obj[ModelName].ToString();
                var displayName = obj[LocalizedDisplayName].ToString();
                var id = obj[CiId].ToString();
                nodes.Add(new Node(id, displayName, parentID, modelName));
            }

            return nodes;
        }

        public void BuildRecursive()
        {
            WriteLine("Building Recursively, This may take time...");
            var sw = Stopwatch.StartNew();
            Connect();
            var nodes = BuildApplicationTreeRecursive(6000, "0", true, false);
            WriteLine($"Folders: {nodes.Count(i => !i.IsApp)}");
            WriteLine($"Applications: {nodes.Count(i => i.IsApp)}");
            WriteLine($"Total: {nodes.Count}\n");

            sw.Stop();
            WriteLine($"Applications Recursively loaded: {sw.Elapsed}");
        }

        public void WriteLine(string msg)
        {
            Console.WriteLine($"{DateTime.Now}: {msg}");
        }
    }

    public class Node
    {
        public string Id { get; set; }
        public string Name { get; set; }
        public string ParentContainerNodeId { get; set; }
        public string AppId { get; set; }
        public bool IsApp => !string.IsNullOrEmpty(AppId);

        public Node(string id, string name, string parentContainer, string appId = null)
        {
            Id = id;
            Name = name;
            ParentContainerNodeId = parentContainer;
            AppId = appId;
        }
    }
}
