using CodeWalker.GameFiles;
using CodeWalker.Utils;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Xml;

namespace CodeWalker.Tools
{
    public sealed class YtypPropExportProgress
    {
        public string Status { get; set; }
        public int Current { get; set; }
        public int Total { get; set; }
    }

    public sealed class YtypPropExportResult
    {
        public string ExportedYtypXmlFileName { get; set; }
        public int TotalTargets { get; set; }
        public int ExportedTargets { get; set; }
        public int ExportedTextures { get; set; }
        public int MissingTextures { get; set; }
        public int MissingArchetypes { get; set; }
        public int MissingResources { get; set; }
        public List<string> Errors { get; } = new List<string>();
    }

    public sealed class YtypPropSelectionItem
    {
        public string Key { get; set; }
        public string Label { get; set; }
        public string MloName { get; set; }
        public int Index { get; set; }
        public int ItemCount { get; set; }
    }

    public sealed class YtypPropSelectionInfo
    {
        public int MloCount { get; set; }
        public string PrimaryMloName { get; set; }
        public List<YtypPropSelectionItem> Rooms { get; } = new List<YtypPropSelectionItem>();
        public List<YtypPropSelectionItem> EntitySets { get; } = new List<YtypPropSelectionItem>();
    }

    public sealed class YtypPropExportSelection
    {
        public bool ImportAllMlo { get; set; } = true;
        public HashSet<string> RoomKeys { get; } = new HashSet<string>(StringComparer.InvariantCultureIgnoreCase);
        public HashSet<string> EntitySetKeys { get; } = new HashSet<string>(StringComparer.InvariantCultureIgnoreCase);
    }

    public sealed class YtypPropExporter
    {
        public const string DrawableFolderName = "Drawable";

        private Dictionary<uint, RpfFileEntry> drawableEntryIndex;
        private Dictionary<uint, RpfFileEntry> yddEntryIndex;
        private Dictionary<uint, RpfFileEntry> ytdEntryIndex;
        private Dictionary<ulong, Texture> textureLookupCache;
        private Dictionary<uint, YtdFile> ytdFileCache;
        private HashSet<uint> missingYtdHashes;

        private sealed class ExportTarget
        {
            public RpfFileEntry Entry { get; set; }
            public List<Archetype> Archetypes { get; } = new List<Archetype>();
            public HashSet<uint> TextureDictHashes { get; } = new HashSet<uint>();
        }

        public static bool SupportsInputPath(string path)
        {
            if (string.IsNullOrWhiteSpace(path))
            {
                return false;
            }

            var name = Path.GetFileName(path);
            if (string.IsNullOrWhiteSpace(name))
            {
                return false;
            }

            if (name.EndsWith(".ytyp", StringComparison.InvariantCultureIgnoreCase))
            {
                return true;
            }

            return name.EndsWith(".ytyp.xml", StringComparison.InvariantCultureIgnoreCase);
        }

        public static string GetSuggestedOutputFolderName(string inputPath, string preferredName = null)
        {
            var name = preferredName;
            if (string.IsNullOrWhiteSpace(name))
            {
                name = Path.GetFileName(inputPath);
            }
            if (string.IsNullOrWhiteSpace(name))
            {
                name = "export";
            }

            if (name.EndsWith(".ytyp.xml", StringComparison.InvariantCultureIgnoreCase))
            {
                name = name.Substring(0, name.Length - ".ytyp.xml".Length);
            }
            else
            {
                name = Path.GetFileNameWithoutExtension(name);
            }

            if (string.IsNullOrWhiteSpace(name))
            {
                name = "export";
            }

            var invalidChars = Path.GetInvalidFileNameChars();
            var clean = new string(name.Select(ch => invalidChars.Contains(ch) ? '_' : ch).ToArray()).Trim();
            clean = string.IsNullOrWhiteSpace(clean) ? "export" : clean;
            return "Mlo Extracted - " + clean;
        }

        public static string GetSuggestedOutputFolderPath(string inputPath, string preferredName = null)
        {
            var baseFolder = Path.GetDirectoryName(Path.GetFullPath(inputPath));
            if (string.IsNullOrWhiteSpace(baseFolder))
            {
                baseFolder = Environment.CurrentDirectory;
            }
            return Path.Combine(baseFolder, GetSuggestedOutputFolderName(inputPath, preferredName));
        }

        public YtypPropSelectionInfo LoadSelectionInfo(string inputPath)
        {
            ValidateInputPath(inputPath);

            var ytyp = LoadYtypForExport(inputPath);
            var mlos = ytyp?.AllArchetypes?.OfType<MloArchetype>().ToArray();
            if ((mlos == null) || (mlos.Length == 0))
            {
                throw new Exception("No MLO archetypes were found in this YTYP.");
            }

            var selectionInfo = new YtypPropSelectionInfo()
            {
                MloCount = mlos.Length
            };

            bool multipleMlos = mlos.Length > 1;
            foreach (var mlo in mlos)
            {
                if (mlo == null)
                {
                    continue;
                }

                var mloName = GetDisplayLabel(mlo.Name, "MLO");
                if (!multipleMlos && string.IsNullOrWhiteSpace(selectionInfo.PrimaryMloName))
                {
                    selectionInfo.PrimaryMloName = mloName;
                }

                if (mlo.rooms != null)
                {
                    for (int i = 0; i < mlo.rooms.Length; i++)
                    {
                        var room = mlo.rooms[i];
                        selectionInfo.Rooms.Add(new YtypPropSelectionItem()
                        {
                            Key = BuildRoomKey(mlo, i),
                            Label = BuildRoomLabel(mloName, room, i, multipleMlos),
                            MloName = mloName,
                            Index = i,
                            ItemCount = room?.AttachedObjects?.Length ?? 0
                        });
                    }
                }

                if (mlo.entitySets != null)
                {
                    for (int i = 0; i < mlo.entitySets.Length; i++)
                    {
                        var entitySet = mlo.entitySets[i];
                        selectionInfo.EntitySets.Add(new YtypPropSelectionItem()
                        {
                            Key = BuildEntitySetKey(mlo, i),
                            Label = BuildEntitySetLabel(mloName, entitySet, i, multipleMlos),
                            MloName = mloName,
                            Index = i,
                            ItemCount = entitySet?.Entities?.Length ?? 0
                        });
                    }
                }
            }

            return selectionInfo;
        }

        public YtypPropExportResult Export(GameFileCache fileCache, string inputPath, string outputFolder, bool includeTextures, YtypPropExportSelection selection = null, Action<YtypPropExportProgress> progress = null, Action<string> status = null, int cacheWaitTimeoutMs = 60000)
        {
            if (fileCache == null)
            {
                throw new ArgumentNullException(nameof(fileCache));
            }
            ValidateInputPath(inputPath);

            Directory.CreateDirectory(outputFolder);
            ResetRuntimeCaches();

            ReportProgress(status, progress, "Waiting for file cache...");
            fileCache = WaitForFileCacheReady(fileCache, cacheWaitTimeoutMs);
            if (fileCache == null)
            {
                throw new Exception("Game file cache isn't ready yet. Please wait for the GTA files to finish loading and try again.");
            }

            ReportProgress(status, progress, "Loading YTYP...");
            var ytyp = LoadYtypForExport(inputPath);
            var mlos = ytyp?.AllArchetypes?.OfType<MloArchetype>().ToArray();
            if ((mlos == null) || (mlos.Length == 0))
            {
                throw new Exception("No MLO archetypes were found in this YTYP.");
            }

            var localArchetypes = BuildArchetypeLookup(ytyp);
            var result = new YtypPropExportResult();
            var targets = new Dictionary<string, ExportTarget>(StringComparer.InvariantCultureIgnoreCase);

            foreach (var mlo in mlos)
            {
                AddYtypPropTargets(GetSelectedMloEntities(mlo, selection), localArchetypes, fileCache, targets, result);
                if (mlo.entitySets != null)
                {
                    for (int i = 0; i < mlo.entitySets.Length; i++)
                    {
                        if (!ShouldExportEntitySet(mlo, i, selection))
                        {
                            continue;
                        }

                        var entitySet = mlo.entitySets[i];
                        AddYtypPropTargets(entitySet?.Entities, localArchetypes, fileCache, targets, result);
                    }
                }
            }

            result.TotalTargets = targets.Count;
            if (targets.Count == 0)
            {
                throw new Exception("No prop resource files were found for the selected rooms or entity sets.");
            }

            ReportProgress(status, progress, "Exporting YTYP XML...");
            result.ExportedYtypXmlFileName = ExportYtypXml(ytyp, inputPath, outputFolder);

            var drawableOutputFolder = GetDrawableOutputFolderPath(outputFolder);
            Directory.CreateDirectory(drawableOutputFolder);

            if (includeTextures)
            {
                ReportProgress(status, progress, "Caching shared texture dictionaries...");
                PreloadTextureDictionaries(fileCache, targets.Values.SelectMany(target => BuildTextureDictHashCandidates(fileCache, target)));
            }

            ReportProgress(status, progress, "Found " + targets.Count.ToString() + " prop files to export.", 0, targets.Count);

            int current = 0;
            foreach (var target in targets.Values.OrderBy(t => t.Entry?.Name, StringComparer.InvariantCultureIgnoreCase))
            {
                current++;
                var entry = target.Entry;
                if (entry == null)
                {
                    result.Errors.Add("A prop export target did not have a valid source entry.");
                    continue;
                }

                ReportProgress(status, progress, "Exporting " + current.ToString() + "/" + targets.Count.ToString() + ": " + entry.Name + "...", current, targets.Count);

                try
                {
                    var data = entry.File?.ExtractFile(entry);
                    if (data == null)
                    {
                        result.Errors.Add("Unable to extract " + entry.Path);
                        continue;
                    }

                    var textureOutputFolder = GetTextureOutputFolderPath(drawableOutputFolder, entry);
                    var xml = MetaXml.GetXml(entry, data, out var newFileName, null);
                    if (string.IsNullOrEmpty(xml) || string.IsNullOrEmpty(newFileName))
                    {
                        result.Errors.Add("Unable to convert " + entry.Path + " to XML.");
                        continue;
                    }

                    if (includeTextures)
                    {
                        Directory.CreateDirectory(textureOutputFolder);
                        xml = RewriteEmbeddedTexturePaths(xml, Path.GetFileName(textureOutputFolder));
                    }

                    File.WriteAllText(Path.Combine(drawableOutputFolder, newFileName), xml);
                    result.ExportedTargets++;

                    if (includeTextures)
                    {
                        ExportTargetTextures(fileCache, target, textureOutputFolder, result);
                    }
                }
                catch (Exception ex)
                {
                    result.Errors.Add(entry.Name + ": " + ex.Message);
                }
            }

            ReportProgress(status, progress, "Export complete. " + result.ExportedTargets.ToString() + "/" + result.TotalTargets.ToString() + " prop files exported.", result.TotalTargets, result.TotalTargets);
            return result;
        }

        private void ReportProgress(Action<string> status, Action<YtypPropExportProgress> progress, string message, int current = 0, int total = 0)
        {
            status?.Invoke(message);
            progress?.Invoke(new YtypPropExportProgress()
            {
                Status = message,
                Current = current,
                Total = total
            });
        }

        private void ValidateInputPath(string inputPath)
        {
            if (string.IsNullOrWhiteSpace(inputPath))
            {
                throw new ArgumentException("Input path is required.", nameof(inputPath));
            }
            if (!File.Exists(inputPath))
            {
                throw new FileNotFoundException("The selected YTYP file could not be found.", inputPath);
            }
            if (!SupportsInputPath(inputPath))
            {
                throw new InvalidOperationException("Only .ytyp and .ytyp.xml inputs are supported.");
            }
        }

        private GameFileCache WaitForFileCacheReady(GameFileCache fileCache, int timeoutMs)
        {
            var stopwatch = Stopwatch.StartNew();
            while (!fileCache.IsInited && (stopwatch.ElapsedMilliseconds < timeoutMs))
            {
                Thread.Sleep(50);
            }

            return fileCache.IsInited ? fileCache : null;
        }

        private YtypFile LoadYtypForExport(string inputPath)
        {
            byte[] data;
            string fileName;
            string filePath;
            if (inputPath.EndsWith(".xml", StringComparison.InvariantCultureIgnoreCase))
            {
                var xmlName = Path.GetFileName(inputPath);
                var xmlNameLower = xmlName.ToLowerInvariant();
                int trimLength = 4;
                var metaFormat = XmlMeta.GetXMLFormat(xmlNameLower, out trimLength);

                fileName = xmlName.Substring(0, xmlName.Length - trimLength);
                filePath = Path.Combine(Path.GetDirectoryName(inputPath), fileName);
                var resourceHintPath = Path.Combine(Path.GetDirectoryName(filePath), Path.GetFileNameWithoutExtension(filePath));

                var doc = new XmlDocument();
                doc.LoadXml(File.ReadAllText(inputPath));
                data = XmlMeta.GetData(doc, metaFormat, resourceHintPath);

                if (data == null)
                {
                    throw new Exception("The XML schema is not supported for YTYP import.");
                }
            }
            else
            {
                fileName = Path.GetFileName(inputPath);
                filePath = inputPath;
                data = File.ReadAllBytes(inputPath);
            }

            if (!fileName.EndsWith(".ytyp", StringComparison.InvariantCultureIgnoreCase))
            {
                throw new Exception("Only YTYP inputs are supported.");
            }

            var entry = CreateFileEntry(fileName, filePath, ref data);
            return RpfFile.GetFile<YtypFile>(entry, data);
        }

        private RpfFileEntry CreateFileEntry(string name, string path, ref byte[] data)
        {
            RpfFileEntry entry = null;
            uint rsc7 = (data?.Length > 4) ? BitConverter.ToUInt32(data, 0) : 0;
            if (rsc7 == 0x37435352)
            {
                entry = RpfFile.CreateResourceFileEntry(ref data, 0);
                data = ResourceBuilder.Decompress(data);
            }
            else
            {
                var binaryEntry = new RpfBinaryFileEntry();
                binaryEntry.FileSize = (uint)(data?.Length ?? 0);
                binaryEntry.FileUncompressedSize = binaryEntry.FileSize;
                entry = binaryEntry;
            }

            entry.Name = name;
            entry.NameLower = name?.ToLowerInvariant();
            entry.NameHash = JenkHash.GenHash(entry.NameLower);
            entry.ShortNameHash = JenkHash.GenHash(Path.GetFileNameWithoutExtension(entry.NameLower));
            entry.Path = path;
            return entry;
        }

        private Dictionary<uint, Archetype> BuildArchetypeLookup(YtypFile ytyp)
        {
            var lookup = new Dictionary<uint, Archetype>();
            if (ytyp?.AllArchetypes == null)
            {
                return lookup;
            }

            foreach (var archetype in ytyp.AllArchetypes)
            {
                if (archetype == null)
                {
                    continue;
                }

                uint hash = archetype.Hash;
                if ((hash != 0) && !lookup.ContainsKey(hash))
                {
                    lookup[hash] = archetype;
                }
            }

            return lookup;
        }

        private IEnumerable<MCEntityDef> GetSelectedMloEntities(MloArchetype mlo, YtypPropExportSelection selection)
        {
            var entities = mlo?.entities;
            if ((entities == null) || (entities.Length == 0))
            {
                return null;
            }

            if ((selection == null) || selection.ImportAllMlo)
            {
                return entities;
            }

            var selectedEntities = new List<MCEntityDef>();
            var selectedIndexes = new HashSet<uint>();
            var rooms = mlo.rooms;
            if (rooms == null)
            {
                return selectedEntities;
            }

            for (int i = 0; i < rooms.Length; i++)
            {
                if (!selection.RoomKeys.Contains(BuildRoomKey(mlo, i)))
                {
                    continue;
                }

                var attachedObjects = rooms[i]?.AttachedObjects;
                if (attachedObjects == null)
                {
                    continue;
                }

                foreach (var objIndex in attachedObjects)
                {
                    if ((objIndex >= entities.Length) || !selectedIndexes.Add(objIndex))
                    {
                        continue;
                    }

                    var entity = entities[objIndex];
                    if (entity != null)
                    {
                        selectedEntities.Add(entity);
                    }
                }
            }

            return selectedEntities;
        }

        private bool ShouldExportEntitySet(MloArchetype mlo, int entitySetIndex, YtypPropExportSelection selection)
        {
            if (selection == null)
            {
                return true;
            }

            return selection.EntitySetKeys.Contains(BuildEntitySetKey(mlo, entitySetIndex));
        }

        private void AddYtypPropTargets(IEnumerable<MCEntityDef> entities, Dictionary<uint, Archetype> localArchetypes, GameFileCache fileCache, Dictionary<string, ExportTarget> targets, YtypPropExportResult result)
        {
            if (entities == null)
            {
                return;
            }

            foreach (var entity in entities)
            {
                if (entity == null)
                {
                    continue;
                }

                uint archetypeHash = entity._Data.archetypeName;
                if (archetypeHash == 0)
                {
                    result.MissingArchetypes++;
                    continue;
                }

                var archetype = ResolveArchetypeForExport(entity, localArchetypes, fileCache);
                if (archetype == null)
                {
                    result.MissingArchetypes++;
                }

                var entry = GetExportEntryForArchetype(fileCache, archetype, archetypeHash);
                if (entry == null)
                {
                    result.MissingResources++;
                    continue;
                }

                var key = entry.Path ?? entry.Name ?? (archetype?.Name ?? archetypeHash.ToString("X8"));
                if (!targets.TryGetValue(key, out var target))
                {
                    target = new ExportTarget() { Entry = entry };
                    targets[key] = target;
                }

                if ((archetype != null) && !target.Archetypes.Any(a => AreEquivalentExportArchetypes(a, archetype)))
                {
                    target.Archetypes.Add(archetype);
                }

                AddTextureDictHashes(target, archetype, archetypeHash);
            }
        }

        private Archetype ResolveArchetypeForExport(MCEntityDef entity, Dictionary<uint, Archetype> localArchetypes, GameFileCache fileCache)
        {
            if (entity == null)
            {
                return null;
            }

            uint hash = entity._Data.archetypeName;
            if ((hash != 0) && localArchetypes.TryGetValue(hash, out var localArchetype))
            {
                return localArchetype;
            }

            return fileCache.GetArchetype(hash);
        }

        private RpfFileEntry GetExportEntryForArchetype(GameFileCache fileCache, Archetype archetype, uint fallbackHash)
        {
            if (archetype != null)
            {
                uint drawableDictHash = archetype.DrawableDict;
                if (drawableDictHash != 0)
                {
                    var yddEntry = fileCache.GetYddEntry(drawableDictHash);
                    if (yddEntry != null)
                    {
                        return yddEntry;
                    }

                    yddEntry = FindEntryByHash(fileCache, drawableDictHash, ".ydd");
                    if (yddEntry != null)
                    {
                        return yddEntry;
                    }
                }
            }

            uint hash = fallbackHash;
            if (archetype != null)
            {
                hash = archetype.Hash;
            }
            if ((hash == 0) && (fallbackHash != 0))
            {
                hash = fallbackHash;
            }
            if (hash == 0)
            {
                return null;
            }

            return GetExportEntryForHash(fileCache, hash);
        }

        private RpfFileEntry GetExportEntryForHash(GameFileCache fileCache, uint hash)
        {
            if (hash == 0)
            {
                return null;
            }

            return fileCache.GetYdrEntry(hash)
                ?? fileCache.GetYftEntry(hash)
                ?? fileCache.GetYddEntry(hash)
                ?? FindEntryByHash(fileCache, hash, ".ydr", ".yft", ".ydd");
        }

        private void AddTextureDictHashes(ExportTarget target, Archetype archetype, uint fallbackHash)
        {
            if (target == null)
            {
                return;
            }

            if (archetype != null)
            {
                if (archetype.TextureDict != 0)
                {
                    target.TextureDictHashes.Add(archetype.TextureDict);
                }
                else if (archetype.Hash != 0)
                {
                    target.TextureDictHashes.Add(archetype.Hash);
                }
            }
            else if (fallbackHash != 0)
            {
                target.TextureDictHashes.Add(fallbackHash);
            }
        }

        private RpfFileEntry FindEntryByHash(GameFileCache fileCache, uint hash, params string[] extensions)
        {
            if ((fileCache?.RpfMan == null) || (hash == 0) || (extensions == null) || (extensions.Length == 0))
            {
                return null;
            }

            if (TryGetIndexedEntryByHash(fileCache, hash, extensions, out var indexedEntry))
            {
                return indexedEntry;
            }

            var extSet = new HashSet<string>(extensions, StringComparer.InvariantCultureIgnoreCase);
            var scanned = new HashSet<string>(StringComparer.InvariantCultureIgnoreCase);

            foreach (var rpf in fileCache.RpfMan.ModRpfs)
            {
                if (rpf == null) continue;
                if (!scanned.Add(rpf.Path ?? string.Empty)) continue;
                var entry = FindEntryByHash(rpf, hash, extSet);
                if (entry != null) return entry;
            }
            foreach (var rpf in fileCache.RpfMan.AllRpfs)
            {
                if (rpf == null) continue;
                if (!scanned.Add(rpf.Path ?? string.Empty)) continue;
                var entry = FindEntryByHash(rpf, hash, extSet);
                if (entry != null) return entry;
            }

            return null;
        }
        private void ResetRuntimeCaches()
        {
            drawableEntryIndex = null;
            yddEntryIndex = null;
            ytdEntryIndex = null;
            textureLookupCache = new Dictionary<ulong, Texture>();
            ytdFileCache = new Dictionary<uint, YtdFile>();
            missingYtdHashes = new HashSet<uint>();
        }
        private bool TryGetIndexedEntryByHash(GameFileCache fileCache, uint hash, string[] extensions, out RpfFileEntry entry)
        {
            entry = null;

            if (IsExtensionSet(extensions, ".ydd"))
            {
                return GetOrBuildEntryIndex(fileCache, ref yddEntryIndex, ".ydd").TryGetValue(hash, out entry);
            }

            if (IsExtensionSet(extensions, ".ytd"))
            {
                return GetOrBuildEntryIndex(fileCache, ref ytdEntryIndex, ".ytd").TryGetValue(hash, out entry);
            }

            if (IsExtensionSet(extensions, ".ydr", ".yft", ".ydd"))
            {
                return GetOrBuildEntryIndex(fileCache, ref drawableEntryIndex, ".ydr", ".yft", ".ydd").TryGetValue(hash, out entry);
            }

            return false;
        }
        private bool IsExtensionSet(string[] actualExtensions, params string[] expectedExtensions)
        {
            if ((actualExtensions == null) || (expectedExtensions == null) || (actualExtensions.Length != expectedExtensions.Length))
            {
                return false;
            }

            var actual = new HashSet<string>(actualExtensions, StringComparer.InvariantCultureIgnoreCase);
            var expected = new HashSet<string>(expectedExtensions, StringComparer.InvariantCultureIgnoreCase);
            return actual.SetEquals(expected);
        }
        private Dictionary<uint, RpfFileEntry> GetOrBuildEntryIndex(GameFileCache fileCache, ref Dictionary<uint, RpfFileEntry> index, params string[] extensions)
        {
            if (index != null)
            {
                return index;
            }

            var built = new Dictionary<uint, RpfFileEntry>();
            var extSet = new HashSet<string>(extensions, StringComparer.InvariantCultureIgnoreCase);
            var scanned = new HashSet<string>(StringComparer.InvariantCultureIgnoreCase);

            AddEntriesToIndex(fileCache?.RpfMan?.ModRpfs, extSet, scanned, built);
            AddEntriesToIndex(fileCache?.RpfMan?.AllRpfs, extSet, scanned, built);

            index = built;
            return index;
        }
        private void AddEntriesToIndex(IEnumerable<RpfFile> rpfs, HashSet<string> extensions, HashSet<string> scanned, Dictionary<uint, RpfFileEntry> index)
        {
            if (rpfs == null)
            {
                return;
            }

            foreach (var rpf in rpfs)
            {
                if (rpf == null) continue;
                if (!scanned.Add(rpf.Path ?? string.Empty)) continue;

                var entries = rpf.AllEntries;
                if (entries == null) continue;

                foreach (var entry in entries)
                {
                    if (!(entry is RpfFileEntry fileEntry))
                    {
                        continue;
                    }

                    var extension = Path.GetExtension(fileEntry.NameLower ?? string.Empty);
                    if (!extensions.Contains(extension))
                    {
                        continue;
                    }

                    AddEntryIndexValue(index, fileEntry.ShortNameHash, fileEntry);
                    AddEntryIndexValue(index, fileEntry.NameHash, fileEntry);
                    if (TryParseHashedFileName(fileEntry.GetShortNameLower(), out var parsedHash))
                    {
                        AddEntryIndexValue(index, parsedHash, fileEntry);
                    }
                }
            }
        }
        private void AddEntryIndexValue(Dictionary<uint, RpfFileEntry> index, uint hash, RpfFileEntry entry)
        {
            if ((hash == 0) || (entry == null))
            {
                return;
            }

            if (!index.ContainsKey(hash))
            {
                index[hash] = entry;
            }
        }

        private RpfFileEntry FindEntryByHash(RpfFile rpf, uint hash, HashSet<string> extensions)
        {
            var entries = rpf?.AllEntries;
            if (entries == null)
            {
                return null;
            }

            foreach (var entry in entries)
            {
                if (!(entry is RpfFileEntry fileEntry))
                {
                    continue;
                }

                var extension = Path.GetExtension(fileEntry.NameLower ?? string.Empty);
                if (!extensions.Contains(extension))
                {
                    continue;
                }

                if ((fileEntry.ShortNameHash == hash) || (fileEntry.NameHash == hash))
                {
                    return fileEntry;
                }

                if (TryParseHashedFileName(fileEntry.GetShortNameLower(), out var parsedHash) && (parsedHash == hash))
                {
                    return fileEntry;
                }
            }

            return null;
        }

        private bool TryParseHashedFileName(string shortName, out uint hash)
        {
            hash = 0;
            if (string.IsNullOrWhiteSpace(shortName))
            {
                return false;
            }

            if (!shortName.StartsWith("hash_", StringComparison.InvariantCultureIgnoreCase))
            {
                return false;
            }

            var hex = shortName.Substring(5);
            if (hex.StartsWith("0x", StringComparison.InvariantCultureIgnoreCase))
            {
                hex = hex.Substring(2);
            }

            return uint.TryParse(hex, NumberStyles.HexNumber, CultureInfo.InvariantCulture, out hash);
        }

        private bool AreEquivalentExportArchetypes(Archetype left, Archetype right)
        {
            if ((left == null) || (right == null))
            {
                return false;
            }

            return ((uint)left.Hash == (uint)right.Hash)
                && ((uint)left.TextureDict == (uint)right.TextureDict)
                && ((uint)left.DrawableDict == (uint)right.DrawableDict);
        }

        private string ExportYtypXml(YtypFile ytyp, string inputPath, string outputFolder)
        {
            var xml = MetaXml.GetXml(ytyp, out var _);
            if (string.IsNullOrWhiteSpace(xml))
            {
                throw new Exception("Unable to convert the source YTYP to XML.");
            }

            var fileName = GetExportedYtypXmlFileName(inputPath);
            File.WriteAllText(Path.Combine(outputFolder, fileName), xml);
            return fileName;
        }

        private string GetExportedYtypXmlFileName(string inputPath)
        {
            var fileName = Path.GetFileName(inputPath);
            if (string.IsNullOrWhiteSpace(fileName))
            {
                return "export.ytyp.xml";
            }

            if (fileName.EndsWith(".ytyp.xml", StringComparison.InvariantCultureIgnoreCase))
            {
                return fileName;
            }

            if (fileName.EndsWith(".ytyp", StringComparison.InvariantCultureIgnoreCase))
            {
                return fileName + ".xml";
            }

            return Path.GetFileNameWithoutExtension(fileName) + ".ytyp.xml";
        }

        private static string GetDrawableOutputFolderPath(string outputFolder)
        {
            return Path.Combine(outputFolder, DrawableFolderName);
        }

        private void PreloadTextureDictionaries(GameFileCache fileCache, IEnumerable<uint> textureDictHashes)
        {
            if (textureDictHashes == null)
            {
                return;
            }

            foreach (var textureDictHash in textureDictHashes.Where(hash => hash != 0).Distinct())
            {
                GetOrLoadYtd(fileCache, textureDictHash);
            }
        }

        private void ExportTargetTextures(GameFileCache fileCache, ExportTarget target, string textureFolder, YtypPropExportResult result)
        {
            var exportFile = LoadTargetFileForTextureExport(fileCache, target.Entry);
            if (exportFile == null)
            {
                result.Errors.Add("Unable to load " + target.Entry?.Name + " for texture export.");
                return;
            }

            Directory.CreateDirectory(textureFolder);

            var textures = new Dictionary<string, Texture>(StringComparer.InvariantCultureIgnoreCase);
            var missingTextures = new HashSet<string>(StringComparer.InvariantCultureIgnoreCase);
            var textureDictHashes = BuildTextureDictHashCandidates(fileCache, target);

            CollectTexturesForFile(fileCache, target, exportFile, textureDictHashes, textures, missingTextures);

            foreach (var texture in textures.Values)
            {
                if (texture == null)
                {
                    continue;
                }

                try
                {
                    var textureName = texture.Name ?? "null";
                    File.WriteAllBytes(Path.Combine(textureFolder, textureName + ".dds"), DDSIO.GetDDSFile(texture));
                    result.ExportedTextures++;
                }
                catch (Exception ex)
                {
                    result.Errors.Add((texture.Name ?? "unknown_texture") + ": " + ex.Message);
                }
            }

            result.MissingTextures += missingTextures.Count;
        }

        private GameFile LoadTargetFileForTextureExport(GameFileCache fileCache, RpfFileEntry entry)
        {
            if (entry == null)
            {
                return null;
            }

            var extension = Path.GetExtension(entry.NameLower ?? string.Empty);
            switch (extension)
            {
                case ".ydr":
                    var ydr = new YdrFile(entry);
                    return fileCache.LoadFile(ydr) ? ydr : null;
                case ".ydd":
                    var ydd = new YddFile(entry);
                    return fileCache.LoadFile(ydd) ? ydd : null;
                case ".yft":
                    var yft = new YftFile(entry);
                    return fileCache.LoadFile(yft) ? yft : null;
                default:
                    return null;
            }
        }

        private string GetTextureOutputFolderPath(string outputFolder, RpfFileEntry entry)
        {
            var folderName = Path.GetFileNameWithoutExtension(entry?.Name ?? "textures");
            if (string.IsNullOrWhiteSpace(folderName))
            {
                folderName = "textures";
            }

            return Path.Combine(outputFolder, folderName);
        }

        private string RewriteEmbeddedTexturePaths(string xml, string textureFolderName)
        {
            if (string.IsNullOrWhiteSpace(xml) || string.IsNullOrWhiteSpace(textureFolderName))
            {
                return xml;
            }

            return Regex.Replace(
                xml,
                @"<FileName>([^<]+\.dds)</FileName>",
                match =>
                {
                    var fileName = match.Groups[1].Value;
                    if (string.IsNullOrWhiteSpace(fileName) || fileName.Contains("\\") || fileName.Contains("/"))
                    {
                        return match.Value;
                    }

                    return "<FileName>" + textureFolderName + "\\" + fileName + "</FileName>";
                },
                RegexOptions.IgnoreCase);
        }

        private IEnumerable<uint> BuildTextureDictHashCandidates(GameFileCache fileCache, ExportTarget target)
        {
            var hashes = new HashSet<uint>();
            if (target?.TextureDictHashes != null)
            {
                foreach (var hash in target.TextureDictHashes)
                {
                    AddTextureDictHashCandidate(fileCache, hashes, hash);
                }
            }

            if ((hashes.Count == 0) && (target?.Archetypes != null))
            {
                foreach (var archetype in target.Archetypes)
                {
                    var hash = (uint)(archetype?.TextureDict ?? 0);
                    if (hash == 0)
                    {
                        hash = (uint)(archetype?.Hash ?? 0);
                    }
                    AddTextureDictHashCandidate(fileCache, hashes, hash);
                }
            }

            if (hashes.Count == 0)
            {
                AddTextureDictHashCandidate(fileCache, hashes, target?.Entry?.ShortNameHash ?? 0);
            }

            return hashes;
        }

        private void AddTextureDictHashCandidate(GameFileCache fileCache, HashSet<uint> hashes, uint hash)
        {
            if ((hash == 0) || (hashes == null))
            {
                return;
            }

            hashes.Add(hash);

            var hdHash = fileCache?.TryGetHDTextureHash(hash) ?? 0;
            if ((hdHash != 0) && (hdHash != hash))
            {
                hashes.Add(hdHash);
            }
        }

        private void CollectTexturesForFile(GameFileCache fileCache, ExportTarget target, GameFile exportFile, IEnumerable<uint> textureDictHashes, Dictionary<string, Texture> textures, HashSet<string> missingTextures)
        {
            var defaultArchetype = GetDefaultArchetype(target);

            if (exportFile is YdrFile ydr)
            {
                var archetypeHash = target?.Entry?.ShortNameHash ?? 0;
                CollectTexturesForDrawable(fileCache, target, archetypeHash, defaultArchetype, textureDictHashes, ydr.Drawable, textures, missingTextures);
                return;
            }

            if (exportFile is YddFile ydd)
            {
                if (ydd.Dict != null)
                {
                    foreach (var kvp in ydd.Dict)
                    {
                        var drawableArchetype = FindTargetArchetype(target, kvp.Key) ?? defaultArchetype;
                        CollectTexturesForDrawable(fileCache, target, kvp.Key, drawableArchetype, textureDictHashes, kvp.Value, textures, missingTextures);
                    }
                }
                else if (ydd.Drawables != null)
                {
                    foreach (var drawable in ydd.Drawables)
                    {
                        CollectTexturesForDrawable(fileCache, target, 0, defaultArchetype, textureDictHashes, drawable, textures, missingTextures);
                    }
                }
                return;
            }

            if (exportFile is YftFile yft)
            {
                var fragment = yft.Fragment;
                if (fragment == null)
                {
                    return;
                }

                uint archetypeHash = target?.Entry?.ShortNameHash ?? 0;
                var yftShortName = exportFile?.RpfFileEntry?.GetShortNameLower();
                if (yftShortName?.EndsWith("_hi") ?? false)
                {
                    archetypeHash = JenkHash.GenHash(yftShortName.Substring(0, yftShortName.Length - 3));
                }

                var fragmentArchetype = FindTargetArchetype(target, archetypeHash) ?? defaultArchetype;

                CollectTexturesForDrawable(fileCache, target, archetypeHash, fragmentArchetype, textureDictHashes, fragment.Drawable, textures, missingTextures);
                CollectTexturesForDrawable(fileCache, target, archetypeHash, fragmentArchetype, textureDictHashes, fragment.DrawableCloth, textures, missingTextures);

                if (fragment.DrawableArray?.data_items != null)
                {
                    foreach (var drawable in fragment.DrawableArray.data_items)
                    {
                        CollectTexturesForDrawable(fileCache, target, archetypeHash, fragmentArchetype, textureDictHashes, drawable, textures, missingTextures);
                    }
                }

                if (fragment.Cloths?.data_items != null)
                {
                    foreach (var cloth in fragment.Cloths.data_items)
                    {
                        CollectTexturesForDrawable(fileCache, target, archetypeHash, fragmentArchetype, textureDictHashes, cloth?.Drawable, textures, missingTextures);
                    }
                }

                var childDrawables = fragment.PhysicsLODGroup?.PhysicsLOD1?.Children?.data_items;
                if (childDrawables != null)
                {
                    foreach (var child in childDrawables)
                    {
                        CollectTexturesForDrawable(fileCache, target, archetypeHash, fragmentArchetype, textureDictHashes, child?.Drawable1, textures, missingTextures);
                        CollectTexturesForDrawable(fileCache, target, archetypeHash, fragmentArchetype, textureDictHashes, child?.Drawable2, textures, missingTextures);
                    }
                }
            }
        }

        private Archetype GetDefaultArchetype(ExportTarget target)
        {
            if (target?.Archetypes == null)
            {
                return null;
            }

            var entryHash = target.Entry?.ShortNameHash ?? 0;
            return FindTargetArchetype(target, entryHash) ?? target.Archetypes.FirstOrDefault();
        }

        private Archetype FindTargetArchetype(ExportTarget target, uint archetypeHash)
        {
            if ((target?.Archetypes == null) || (archetypeHash == 0))
            {
                return null;
            }

            foreach (var archetype in target.Archetypes)
            {
                if (archetype == null)
                {
                    continue;
                }

                if (((uint)archetype.Hash == archetypeHash) || ((uint)archetype.DrawableDict == archetypeHash))
                {
                    return archetype;
                }
            }

            return null;
        }

        private void CollectTexturesForDrawable(GameFileCache fileCache, ExportTarget target, uint archetypeHashHint, Archetype fallbackArchetype, IEnumerable<uint> textureDictHashes, DrawableBase drawable, Dictionary<string, Texture> textures, HashSet<string> missingTextures)
        {
            if (drawable == null)
            {
                return;
            }

            var drawableArchetypeHash = GetDrawableArchetypeHash(drawable);
            if (drawableArchetypeHash == 0)
            {
                drawableArchetypeHash = archetypeHashHint;
            }

            var drawableArchetype = FindTargetArchetype(target, drawableArchetypeHash) ?? fileCache.GetArchetype(drawableArchetypeHash) ?? fallbackArchetype;
            var primaryTextureDictHash = (uint)(drawableArchetype?.TextureDict ?? 0);
            if (primaryTextureDictHash == 0)
            {
                primaryTextureDictHash = drawableArchetypeHash;
            }

            var orderedTextureDictHashes = BuildOrderedTextureDictHashes(fileCache, primaryTextureDictHash, textureDictHashes);

            if (drawable.ShaderGroup?.TextureDictionary?.Textures?.data_items != null)
            {
                foreach (var texture in drawable.ShaderGroup.TextureDictionary.Textures.data_items)
                {
                    AddTextureToExport(textures, texture);
                }
            }

            if (drawable.ShaderGroup?.Shaders?.data_items == null)
            {
                return;
            }

            foreach (var shader in drawable.ShaderGroup.Shaders.data_items)
            {
                var parameters = shader?.ParametersList?.Parameters;
                if (parameters == null)
                {
                    continue;
                }

                foreach (var parameter in parameters)
                {
                    var textureRef = parameter.Data as TextureBase;
                    if (textureRef == null)
                    {
                        continue;
                    }

                    if (textureRef is Texture embeddedTexture)
                    {
                        AddTextureToExport(textures, embeddedTexture);
                        continue;
                    }

                    var resolvedTexture = ResolveTextureReference(fileCache, textureRef.NameHash, orderedTextureDictHashes);
                    if (resolvedTexture != null)
                    {
                        AddTextureToExport(textures, resolvedTexture);
                    }
                    else
                    {
                        var missingName = string.IsNullOrEmpty(textureRef.Name) ? textureRef.NameHash.ToString() : textureRef.Name;
                        missingTextures.Add(missingName);
                    }
                }
            }
        }

        private uint GetDrawableArchetypeHash(DrawableBase drawable)
        {
            if (drawable is Drawable dwbl)
            {
                var drawableName = dwbl.Name?.ToLowerInvariant();
                if (!string.IsNullOrWhiteSpace(drawableName))
                {
                    drawableName = drawableName.Replace(".#dr", string.Empty).Replace(".#dd", string.Empty);
                    return JenkHash.GenHash(drawableName);
                }
            }
            else if (drawable is FragDrawable fragDrawable)
            {
                var yft = fragDrawable.Owner as YftFile;
                var entryHash = (uint)(yft?.RpfFileEntry?.ShortNameHash ?? 0);
                var shortName = yft?.RpfFileEntry?.GetShortNameLower();
                if (shortName?.EndsWith("_hi") ?? false)
                {
                    entryHash = JenkHash.GenHash(shortName.Substring(0, shortName.Length - 3));
                }
                return entryHash;
            }

            return 0;
        }

        private IEnumerable<uint> BuildOrderedTextureDictHashes(GameFileCache fileCache, uint primaryTextureDictHash, IEnumerable<uint> sharedTextureDictHashes)
        {
            var ordered = new List<uint>();
            var seen = new HashSet<uint>();

            AddOrderedTextureDictHash(fileCache, ordered, seen, primaryTextureDictHash);

            if (sharedTextureDictHashes != null)
            {
                foreach (var hash in sharedTextureDictHashes)
                {
                    AddOrderedTextureDictHash(fileCache, ordered, seen, hash);
                }
            }

            return ordered;
        }

        private void AddOrderedTextureDictHash(GameFileCache fileCache, List<uint> ordered, HashSet<uint> seen, uint hash)
        {
            if ((hash == 0) || (ordered == null) || (seen == null))
            {
                return;
            }

            if (seen.Add(hash))
            {
                ordered.Add(hash);
            }

            var hdHash = fileCache?.TryGetHDTextureHash(hash) ?? 0;
            if ((hdHash != 0) && (hdHash != hash) && seen.Add(hdHash))
            {
                ordered.Add(hdHash);
            }
        }

        private void AddTextureToExport(Dictionary<string, Texture> textures, Texture texture)
        {
            if (texture == null)
            {
                return;
            }

            textures[texture.Name ?? "null"] = texture;
        }

        private Texture ResolveTextureReference(GameFileCache fileCache, uint textureHash, IEnumerable<uint> textureDictHashes)
        {
            var visitedYtds = new HashSet<uint>();
            if (textureDictHashes != null)
            {
                foreach (var textureDictHash in textureDictHashes)
                {
                    var texture = ResolveTextureReference(fileCache, textureHash, textureDictHash, visitedYtds);
                    if (texture != null)
                    {
                        return texture;
                    }
                }
            }

            var ytd = fileCache.TryGetTextureDictForTexture(textureHash);
            if ((ytd != null) && !ytd.Loaded)
            {
                fileCache.LoadFile(ytd);
            }

            return ytd?.TextureDict?.Lookup(textureHash);
        }

        private Texture ResolveTextureReference(GameFileCache fileCache, uint textureHash, uint textureDictHash, HashSet<uint> visitedYtds)
        {
            while ((textureDictHash != 0) && visitedYtds.Add(textureDictHash))
            {
                var texture = TryGetTextureFromYtd(fileCache, textureHash, textureDictHash);
                if (texture != null)
                {
                    return texture;
                }

                textureDictHash = fileCache.TryGetParentYtdHash(textureDictHash);
            }

            return null;
        }

        private Texture TryGetTextureFromYtd(GameFileCache fileCache, uint textureHash, uint textureDictHash)
        {
            if (textureDictHash == 0)
            {
                return null;
            }

            ulong cacheKey = (((ulong)textureDictHash) << 32) | textureHash;
            if ((textureLookupCache != null) && textureLookupCache.TryGetValue(cacheKey, out var cachedTexture))
            {
                return cachedTexture;
            }

            var ytd = GetOrLoadYtd(fileCache, textureDictHash);
            if (ytd == null)
            {
                if (textureLookupCache != null)
                {
                    textureLookupCache[cacheKey] = null;
                }
                return null;
            }

            var texture = ytd.Loaded ? ytd.TextureDict?.Lookup(textureHash) : null;
            if (textureLookupCache != null)
            {
                textureLookupCache[cacheKey] = texture;
            }
            return texture;
        }

        private YtdFile GetOrLoadYtd(GameFileCache fileCache, uint textureDictHash)
        {
            if (textureDictHash == 0)
            {
                return null;
            }

            if ((missingYtdHashes != null) && missingYtdHashes.Contains(textureDictHash))
            {
                return null;
            }

            if ((ytdFileCache != null) && ytdFileCache.TryGetValue(textureDictHash, out var cachedYtd))
            {
                return EnsureYtdLoaded(fileCache, cachedYtd) ? cachedYtd : null;
            }

            var ytd = fileCache.GetYtd(textureDictHash);
            if (ytd == null)
            {
                var ytdEntry = FindEntryByHash(fileCache, textureDictHash, ".ytd");
                if (ytdEntry != null)
                {
                    ytd = new YtdFile(ytdEntry);
                }
            }

            if (ytd == null)
            {
                missingYtdHashes?.Add(textureDictHash);
                return null;
            }

            if (!EnsureYtdLoaded(fileCache, ytd))
            {
                missingYtdHashes?.Add(textureDictHash);
                return null;
            }

            ytdFileCache[textureDictHash] = ytd;
            return ytd;
        }

        private bool EnsureYtdLoaded(GameFileCache fileCache, YtdFile ytd)
        {
            if (ytd == null)
            {
                return false;
            }

            if (!ytd.Loaded)
            {
                fileCache.LoadFile(ytd);
            }

            int tries = 0;
            while (!ytd.Loaded && (tries < 500))
            {
                Thread.Sleep(10);
                tries++;
            }

            return ytd.Loaded;
        }

        private static string BuildRoomKey(MloArchetype mlo, int roomIndex)
        {
            return string.Format(CultureInfo.InvariantCulture, "room:{0:X8}:{1}", mlo?.Hash ?? 0, roomIndex);
        }

        private static string BuildEntitySetKey(MloArchetype mlo, int entitySetIndex)
        {
            return string.Format(CultureInfo.InvariantCulture, "entityset:{0:X8}:{1}", mlo?.Hash ?? 0, entitySetIndex);
        }

        private string BuildRoomLabel(string mloName, MCMloRoomDef room, int roomIndex, bool includeMloName)
        {
            var roomName = GetDisplayLabel(room?.RoomName, "Room " + roomIndex.ToString(CultureInfo.InvariantCulture));
            var objectCount = room?.AttachedObjects?.Length ?? 0;
            var label = roomName + " (" + objectCount.ToString(CultureInfo.InvariantCulture) + ")";
            return includeMloName ? (mloName + " - " + label) : label;
        }

        private string BuildEntitySetLabel(string mloName, MCMloEntitySet entitySet, int entitySetIndex, bool includeMloName)
        {
            var entitySetName = GetDisplayLabel(entitySet?.Name, "Entity Set " + entitySetIndex.ToString(CultureInfo.InvariantCulture));
            var objectCount = entitySet?.Entities?.Length ?? 0;
            var label = entitySetName + " (" + objectCount.ToString(CultureInfo.InvariantCulture) + ")";
            return includeMloName ? (mloName + " - " + label) : label;
        }

        private string GetDisplayLabel(string value, string fallback)
        {
            return string.IsNullOrWhiteSpace(value) ? fallback : value.Trim();
        }
    }
}
