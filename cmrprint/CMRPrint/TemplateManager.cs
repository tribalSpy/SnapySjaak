using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Xml.Serialization;

namespace CMRPrint
{
    [Serializable]
    public class TemplateIntSetting
    {
        public string FieldName { get; set; } = string.Empty;
        public int Value { get; set; }
    }

    [Serializable]
    public class TemplatePointSetting
    {
        public string FieldName { get; set; } = string.Empty;
        public float X { get; set; }
        public float Y { get; set; }
    }

    [Serializable]
    public class CmrTemplate
    {
        public string Name { get; set; } = string.Empty;
        [XmlIgnore]
        public Dictionary<string, int> FontSizes { get; set; } = new();
        [XmlIgnore]
        public Dictionary<string, int> VerticalOffsets { get; set; } = new();
        [XmlIgnore]
        public Dictionary<string, PointF> FieldPositions { get; set; } = new();
        [XmlIgnore]
        public Dictionary<string, int> FieldWidths { get; set; } = new();
        [XmlIgnore]
        public Dictionary<string, int> FieldHeights { get; set; } = new();
        public DateTime CreatedDate { get; set; } = DateTime.Now;
        public List<TemplateIntSetting> FontSizeEntries { get; set; } = new();
        public List<TemplateIntSetting> VerticalOffsetEntries { get; set; } = new();
        public List<TemplatePointSetting> FieldPositionEntries { get; set; } = new();
        public List<TemplateIntSetting> FieldWidthEntries { get; set; } = new();
        public List<TemplateIntSetting> FieldHeightEntries { get; set; } = new();

        public void Save(string filePath)
        {
            SyncSerializableState();
            var serializer = new XmlSerializer(typeof(CmrTemplate));
            using var writer = new StreamWriter(filePath);
            serializer.Serialize(writer, this);
        }

        public static CmrTemplate Load(string filePath)
        {
            var serializer = new XmlSerializer(typeof(CmrTemplate));
            using var reader = new StreamReader(filePath);
            var template = (CmrTemplate?)serializer.Deserialize(reader) ?? new CmrTemplate();
            template.RestoreRuntimeState();
            return template;
        }

        private void SyncSerializableState()
        {
            FontSizeEntries = FontSizes
                .Select(entry => new TemplateIntSetting { FieldName = entry.Key, Value = entry.Value })
                .ToList();
            VerticalOffsetEntries = VerticalOffsets
                .Select(entry => new TemplateIntSetting { FieldName = entry.Key, Value = entry.Value })
                .ToList();
            FieldPositionEntries = FieldPositions
                .Select(entry => new TemplatePointSetting { FieldName = entry.Key, X = entry.Value.X, Y = entry.Value.Y })
                .ToList();
            FieldWidthEntries = FieldWidths
                .Select(entry => new TemplateIntSetting { FieldName = entry.Key, Value = entry.Value })
                .ToList();
            FieldHeightEntries = FieldHeights
                .Select(entry => new TemplateIntSetting { FieldName = entry.Key, Value = entry.Value })
                .ToList();
        }

        private void RestoreRuntimeState()
        {
            FontSizes = FontSizeEntries.ToDictionary(entry => entry.FieldName, entry => entry.Value);
            VerticalOffsets = VerticalOffsetEntries.ToDictionary(entry => entry.FieldName, entry => entry.Value);
            FieldPositions = FieldPositionEntries.ToDictionary(entry => entry.FieldName, entry => new PointF(entry.X, entry.Y));
            FieldWidths = FieldWidthEntries.ToDictionary(entry => entry.FieldName, entry => entry.Value);
            FieldHeights = FieldHeightEntries.ToDictionary(entry => entry.FieldName, entry => entry.Value);
        }
    }

    public static class TemplateManager
    {
        private static readonly string TemplateFolder = Path.Combine(AppDataManager.DataFolder, "Templates");
        private static readonly string LegacyTemplateFolder = Path.Combine(AppDataManager.LegacyFolder, "Templates");

        static TemplateManager()
        {
            EnsureLocalTemplatesInitialized();
            Directory.CreateDirectory(TemplateFolder);
        }

        public static void SaveTemplate(CmrTemplate template)
        {
            var filePath = Path.Combine(TemplateFolder, template.Name + ".xml");
            template.Save(filePath);
        }

        public static CmrTemplate LoadTemplate(string templateName)
        {
            var filePath = Path.Combine(TemplateFolder, templateName + ".xml");
            return File.Exists(filePath) ? CmrTemplate.Load(filePath) : new CmrTemplate { Name = templateName };
        }

        public static List<string> GetAvailableTemplates()
        {
            var templates = new List<string>();
            if (Directory.Exists(TemplateFolder))
            {
                foreach (var file in Directory.GetFiles(TemplateFolder, "*.xml"))
                {
                    templates.Add(Path.GetFileNameWithoutExtension(file));
                }
            }
            return templates;
        }

        public static void DeleteTemplate(string templateName)
        {
            var filePath = Path.Combine(TemplateFolder, templateName + ".xml");
            if (File.Exists(filePath))
            {
                File.Delete(filePath);
            }
        }

        private static void EnsureLocalTemplatesInitialized()
        {
            Directory.CreateDirectory(TemplateFolder);

            var hasLocalTemplates = Directory.Exists(TemplateFolder) && Directory.GetFiles(TemplateFolder, "*.xml").Length > 0;
            if (hasLocalTemplates || !Directory.Exists(LegacyTemplateFolder))
                return;

            foreach (var legacyFile in Directory.GetFiles(LegacyTemplateFolder, "*.xml"))
            {
                var targetFile = Path.Combine(TemplateFolder, Path.GetFileName(legacyFile));
                if (!File.Exists(targetFile))
                {
                    File.Copy(legacyFile, targetFile, overwrite: false);
                }
            }
        }
    }
}
