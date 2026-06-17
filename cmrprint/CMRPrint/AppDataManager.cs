using System;
using System.Collections.Generic;
using System.IO;
using System.Xml.Serialization;

namespace CMRPrint
{
    [Serializable]
    public class AppDataStore
    {
        public List<Customer> Customers { get; set; } = new();
        public List<ProfileRecord> Exporters { get; set; } = new();
        public List<ProfileRecord> TransportInfos { get; set; } = new();
        public List<ProfileRecord> LoadingPlaces { get; set; } = new();
    }

    public static class AppDataManager
    {
        private static readonly string LocalDataFolder = Path.Combine(AppContext.BaseDirectory, "Data");
        private static readonly string LegacyDataFolder = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
            "CMRPrint");
        private static readonly string LocalDataFilePath = Path.Combine(LocalDataFolder, "app-data.xml");
        private static readonly string LegacyDataFilePath = Path.Combine(LegacyDataFolder, "app-data.xml");

        public static string DataFolder => LocalDataFolder;
        public static string LegacyFolder => LegacyDataFolder;
        public static string DataFilePath => LocalDataFilePath;

        static AppDataManager()
        {
            EnsureLocalDataInitialized();
        }

        public static AppDataStore Load()
        {
            Directory.CreateDirectory(LocalDataFolder);

            if (!File.Exists(LocalDataFilePath))
                return new AppDataStore();

            var serializer = new XmlSerializer(typeof(AppDataStore));
            using var stream = File.OpenRead(LocalDataFilePath);
            return (AppDataStore?)serializer.Deserialize(stream) ?? new AppDataStore();
        }

        public static void Save(AppDataStore data)
        {
            Directory.CreateDirectory(LocalDataFolder);

            var serializer = new XmlSerializer(typeof(AppDataStore));
            using var stream = File.Create(LocalDataFilePath);
            serializer.Serialize(stream, data);
        }

        private static void EnsureLocalDataInitialized()
        {
            Directory.CreateDirectory(LocalDataFolder);

            if (File.Exists(LocalDataFilePath) || !File.Exists(LegacyDataFilePath))
                return;

            File.Copy(LegacyDataFilePath, LocalDataFilePath, overwrite: false);
        }
    }
}
