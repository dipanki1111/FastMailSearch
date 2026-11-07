using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using FastMailSearch.Core.Interfaces;
using FastMailSearch.Core.Models;
using FastMailSearch.Infrastructure.Logging;
using OutlookInterop = Microsoft.Office.Interop.Outlook;

namespace FastMailSearch.Infrastructure.Outlook
{
    public class StaWorkerThread : IDisposable
    {
        private readonly BlockingCollection<Action> _queue = new BlockingCollection<Action>();
        private readonly Thread _thread;
        private volatile bool _isDisposed;
        private readonly object _disposeLock = new object();

        public StaWorkerThread()
        {
            _thread = new Thread(RunLoop)
            {
                IsBackground = true,
                Name = "OutlookStaWorker"
            };
            _thread.SetApartmentState(ApartmentState.STA);
            _thread.Start();
        }

        private void RunLoop()
        {
            foreach (var workItem in _queue.GetConsumingEnumerable())
            {
                try
                {
                    workItem();
                }
                catch { }
            }
        }

        public void Invoke(Action action)
        {
            if (action == null) throw new ArgumentNullException(nameof(action));

            if (Thread.CurrentThread == _thread)
            {
                action();
                return;
            }

            if (_isDisposed || _queue.IsAddingCompleted)
            {
                throw new ObjectDisposedException(nameof(StaWorkerThread), "Cannot execute operation because the STA worker thread has been shut down or disposed.");
            }

            var tcs = new TaskCompletionSource<bool>(TaskCreationOptions.RunContinuationsAsynchronously);
            try
            {
                _queue.Add(() =>
                {
                    try
                    {
                        action();
                        tcs.SetResult(true);
                    }
                    catch (Exception ex)
                    {
                        tcs.SetException(ex);
                    }
                });
                tcs.Task.GetAwaiter().GetResult();
            }
            catch (InvalidOperationException ex)
            {
                throw new ObjectDisposedException(nameof(StaWorkerThread), "Cannot execute operation because the STA worker queue is completed: " + ex.Message);
            }
        }

        public T Invoke<T>(Func<T> func)
        {
            if (func == null) throw new ArgumentNullException(nameof(func));

            if (Thread.CurrentThread == _thread)
            {
                return func();
            }

            T result = default!;
            Invoke(() => { result = func(); });
            return result;
        }

        public bool TryInvoke(Action action, TimeSpan timeout)
        {
            if (action == null) throw new ArgumentNullException(nameof(action));

            if (Thread.CurrentThread == _thread)
            {
                action();
                return true;
            }

            if (_isDisposed || _queue.IsAddingCompleted)
            {
                return false;
            }

            var tcs = new TaskCompletionSource<bool>(TaskCreationOptions.RunContinuationsAsynchronously);
            try
            {
                _queue.Add(() =>
                {
                    try
                    {
                        action();
                        tcs.SetResult(true);
                    }
                    catch (Exception ex)
                    {
                        tcs.SetException(ex);
                    }
                });
                return tcs.Task.Wait((int)timeout.TotalMilliseconds);
            }
            catch (InvalidOperationException)
            {
                return false;
            }
        }

        public bool TryInvoke<T>(Func<T> func, TimeSpan timeout, out T result)
        {
            if (func == null) throw new ArgumentNullException(nameof(func));

            if (Thread.CurrentThread == _thread)
            {
                result = func();
                return true;
            }

            if (_isDisposed || _queue.IsAddingCompleted)
            {
                result = default!;
                return false;
            }

            T res = default!;
            bool ok = TryInvoke(() => { res = func(); }, timeout);
            result = res;
            return ok;
        }

        public void Dispose()
        {
            Dispose(3000);
        }

        public void Dispose(int timeoutMs)
        {
            lock (_disposeLock)
            {
                if (_isDisposed) return;
                _isDisposed = true;
                try
                {
                    _queue.CompleteAdding();
                }
                catch (InvalidOperationException) { }
            }

            if (Thread.CurrentThread != _thread && _thread.IsAlive)
            {
                if (timeoutMs > 0)
                {
                    _thread.Join(timeoutMs);
                }
            }
        }
    }

    public class OutlookAdapter : IOutlookAdapter, IFolderBatchReader, IDisposable
    {
        public sealed class FolderEventSubscription
        {
            public OutlookInterop.Items Items { get; }
            public OutlookInterop.ItemsEvents_ItemAddEventHandler? ItemAddHandler { get; }
            public OutlookInterop.ItemsEvents_ItemChangeEventHandler? ItemChangeHandler { get; }
            public OutlookInterop.ItemsEvents_ItemRemoveEventHandler? ItemRemoveHandler { get; }
            public bool IsUnsubscribed { get; private set; }

            public FolderEventSubscription(
                OutlookInterop.Items items,
                OutlookInterop.ItemsEvents_ItemAddEventHandler? itemAddHandler,
                OutlookInterop.ItemsEvents_ItemChangeEventHandler? itemChangeHandler,
                OutlookInterop.ItemsEvents_ItemRemoveEventHandler? itemRemoveHandler = null)
            {
                Items = items;
                ItemAddHandler = itemAddHandler;
                ItemChangeHandler = itemChangeHandler;
                ItemRemoveHandler = itemRemoveHandler;
            }

            public void UnsubscribeAndRelease()
            {
                if (IsUnsubscribed) return;
                IsUnsubscribed = true;

                try
                {
                    if (ItemAddHandler != null)
                    {
                        Items.ItemAdd -= ItemAddHandler;
                    }
                }
                catch { }

                try
                {
                    if (ItemChangeHandler != null)
                    {
                        Items.ItemChange -= ItemChangeHandler;
                    }
                }
                catch { }

                try
                {
                    if (ItemRemoveHandler != null)
                    {
                        Items.ItemRemove -= ItemRemoveHandler;
                    }
                }
                catch { }

                var tmp = Items;
                ComReleaseHelper.Release(ref tmp!);
            }
        }

        public static bool SafeSubscribeItems(
            OutlookInterop.Items items,
            OutlookInterop.ItemsEvents_ItemAddEventHandler addHandler,
            OutlookInterop.ItemsEvents_ItemChangeEventHandler changeHandler,
            out FolderEventSubscription? subscription)
        {
            subscription = null;
            if (items == null) return false;

            try
            {
                items.ItemAdd += addHandler;
                try
                {
                    items.ItemChange += changeHandler;
                    subscription = new FolderEventSubscription(items, addHandler, changeHandler);
                    return true;
                }
                catch (Exception ex)
                {
                    Logger.Error("[OUTLOOK EVENT ERROR] Failed to subscribe ItemChange; unhooking ItemAdd and releasing items COM object", ex);
                    try
                    {
                        items.ItemAdd -= addHandler;
                    }
                    catch { }
                    var tmp = items;
                    ComReleaseHelper.Release(ref tmp!);
                    return false;
                }
            }
            catch (Exception ex)
            {
                Logger.Error("[OUTLOOK EVENT ERROR] Failed to subscribe ItemAdd; releasing items COM object", ex);
                var tmp = items;
                ComReleaseHelper.Release(ref tmp!);
                return false;
            }
        }

        private OutlookInterop.Application? _outlookApp;
        private readonly StaWorkerThread _staThread = new StaWorkerThread();
        private readonly List<FolderEventSubscription> _subscribedItemsList = new List<FolderEventSubscription>();
        private Action<string, string, string>? _onItemChangedCallback;
        private Action<string, string>? _onItemRemovedCallback;
        private readonly HashSet<string> _outboxFolderEntryIds = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        private readonly HashSet<string> _sentFolderEntryIds = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        public OutlookAdapter(OutlookInterop.Application? outlookApp = null)
        {
            _outlookApp = outlookApp;
        }

        public bool IsOutboxFolder(string storeId, string folderEntryId)
        {
            if (string.IsNullOrEmpty(folderEntryId)) return false;
            lock (_outboxFolderEntryIds)
            {
                return _outboxFolderEntryIds.Contains(folderEntryId);
            }
        }

        public void RegisterOutboxFolderEntryId(string entryId)
        {
            if (!string.IsNullOrEmpty(entryId))
            {
                lock (_outboxFolderEntryIds)
                {
                    _outboxFolderEntryIds.Add(entryId);
                }
            }
        }

        public void RegisterSentFolderEntryId(string entryId)
        {
            if (!string.IsNullOrEmpty(entryId))
            {
                lock (_sentFolderEntryIds)
                {
                    _sentFolderEntryIds.Add(entryId);
                }
            }
        }

        public bool IsSentItemsFolderId(string folderEntryId)
        {
            if (string.IsNullOrEmpty(folderEntryId)) return false;
            lock (_sentFolderEntryIds)
            {
                return _sentFolderEntryIds.Contains(folderEntryId);
            }
        }

        private void DiscoverDefaultFolders(OutlookInterop.NameSpace ns, OutlookInterop.Store? store)
        {
            if (store != null)
            {
                try
                {
                    var outbox = store.GetDefaultFolder(OutlookInterop.OlDefaultFolders.olFolderOutbox);
                    if (outbox != null)
                    {
                        RegisterOutboxFolderEntryId(outbox.EntryID);
                        ComReleaseHelper.Release(ref outbox);
                    }
                }
                catch { }

                try
                {
                    var sent = store.GetDefaultFolder(OutlookInterop.OlDefaultFolders.olFolderSentMail);
                    if (sent != null)
                    {
                        RegisterSentFolderEntryId(sent.EntryID);
                        ComReleaseHelper.Release(ref sent);
                    }
                }
                catch { }
            }
            else if (ns != null)
            {
                try
                {
                    var outbox = ns.GetDefaultFolder(OutlookInterop.OlDefaultFolders.olFolderOutbox);
                    if (outbox != null)
                    {
                        RegisterOutboxFolderEntryId(outbox.EntryID);
                        ComReleaseHelper.Release(ref outbox);
                    }
                }
                catch { }

                try
                {
                    var sent = ns.GetDefaultFolder(OutlookInterop.OlDefaultFolders.olFolderSentMail);
                    if (sent != null)
                    {
                        RegisterSentFolderEntryId(sent.EntryID);
                        ComReleaseHelper.Release(ref sent);
                    }
                }
                catch { }
            }
        }

        private OutlookInterop.Application? GetApplication()
        {
            if (_outlookApp != null)
            {
                try
                {
                    // Verify cached COM proxy is still alive and responsive
                    _ = _outlookApp.Version;
                    return _outlookApp;
                }
                catch
                {
                    // Outlook closed or COM connection disconnected
                    _outlookApp = null;
                }
            }

            // PASSIVE ATTACHMENT ONLY:
            // FastMailSearch must never implicitly launch or instantiate Outlook when Outlook is closed.
            // 1. Quick process check: If OUTLOOK process is not running, return null immediately.
            try
            {
                var procs = System.Diagnostics.Process.GetProcessesByName("OUTLOOK");
                if (procs == null || procs.Length == 0)
                {
                    return null;
                }

                // 2. Passive COM attachment to the running Outlook instance via Running Object Table (ROT).
                _outlookApp = (OutlookInterop.Application)GetActiveComObject("Outlook.Application");
                return _outlookApp;
            }
            catch (Exception ex)
            {
                // Outlook is closing, busy, starting up, or Running Object Table registration is not yet ready.
                // Do NOT invoke new OutlookInterop.Application() under any circumstances.
                Logger.Info($"[OUTLOOK COM] Passive attachment unavailable: {ex.Message}");
                _outlookApp = null;
                return null;
            }
        }

        private static object GetActiveComObject(string progId)
        {
#if NETFRAMEWORK
            return Marshal.GetActiveObject(progId);
#else
            CLSIDFromProgID(progId, out Guid clsid);
            GetActiveObject(ref clsid, IntPtr.Zero, out object obj);
            return obj;
#endif
        }

        [DllImport("oleaut32.dll", PreserveSig = false)]
        private static extern void GetActiveObject(ref Guid rclsid, IntPtr pvReserved, [MarshalAs(UnmanagedType.IUnknown)] out object ppunk);

        [DllImport("ole32.dll")]
        private static extern int CLSIDFromProgID([MarshalAs(UnmanagedType.LPWStr)] string lpszProgID, out Guid pclsid);

        public bool IsOutlookAvailable()
        {
            try
            {
                return _staThread.Invoke(() =>
                {
                    try
                    {
                        var app = GetApplication();
                        if (app == null) return false;
                        string ver = app.Version;
                        return !string.IsNullOrEmpty(ver);
                    }
                    catch
                    {
                        return false;
                    }
                });
            }
            catch (ObjectDisposedException)
            {
                return false;
            }
        }

        public string GetOutlookVersion()
        {
            try
            {
                return _staThread.Invoke(() =>
                {
                    try
                    {
                        var app = GetApplication();
                        return app?.Version ?? "Unavailable";
                    }
                    catch
                    {
                        return "Unavailable";
                    }
                });
            }
            catch (ObjectDisposedException)
            {
                return "Unavailable";
            }
        }

        public string GetOfficeBitness()
        {
            return Environment.Is64BitProcess ? "64-bit" : "32-bit";
        }

        public List<MailboxInfo> GetMailboxes()
        {
            try
            {
                return _staThread.Invoke(() =>
            {
                var result = new List<MailboxInfo>();
                var app = GetApplication();
                if (app == null) return result;

                OutlookInterop.NameSpace? ns = null;
                OutlookInterop.Stores? stores = null;

                try
                {
                    ns = app.GetNamespace("MAPI");
                    stores = ns.Stores;
                    int storeCount = stores?.Count ?? 0;
                    Logger.Info($"[STORE DISCOVERY] Querying Outlook stores (Count={storeCount})...");

                    for (int i = 1; i <= storeCount; i++)
                    {
                        OutlookInterop.Store? store = null;
                        try
                        {
                            store = stores[i];
                            string sId = string.Empty;
                            try { sId = store.StoreID ?? string.Empty; } catch (Exception ex) { Logger.Warn($"Failed to read StoreID at index {i}: {ex.Message}"); }

                            string dispName = string.Empty;
                            try { dispName = store.DisplayName ?? string.Empty; } catch { dispName = $"Store {i}"; }

                            string filePath = string.Empty;
                            try { filePath = store.FilePath ?? string.Empty; } catch { }

                            bool isPst = false;
                            try
                            {
                                isPst = store.ExchangeStoreType == OutlookInterop.OlExchangeStoreType.olNotExchange;
                            }
                            catch
                            {
                                isPst = filePath.EndsWith(".pst", StringComparison.OrdinalIgnoreCase) ||
                                        filePath.EndsWith(".ost", StringComparison.OrdinalIgnoreCase);
                            }

                            if (string.IsNullOrEmpty(sId))
                            {
                                Logger.Warn($"[STORE DISCOVERY] Store index {i} '{dispName}' has empty StoreID, skipping.");
                                continue;
                            }

                            result.Add(new MailboxInfo
                            {
                                StoreId = sId,
                                DisplayName = dispName,
                                FilePath = filePath,
                                IsPst = isPst,
                                IsSelected = true
                            });

                            Logger.Info($"[STORE DISCOVERY] Found store [{i}]: DisplayName='{dispName}', FilePath='{filePath}', IsPst={isPst}, StoreId='{sId}'");
                        }
                        catch (Exception ex)
                        {
                            Logger.Warn($"Failed to read store info at index {i}: {ex.Message}");
                        }
                        finally
                        {
                            ComReleaseHelper.Release(ref store);
                        }
                    }

                    // Fallback 1: If ns.Stores returned nothing, check ns.DefaultStore
                    if (result.Count == 0)
                    {
                        OutlookInterop.Store? defaultStore = null;
                        try
                        {
                            defaultStore = ns.DefaultStore;
                            if (defaultStore != null && !string.IsNullOrEmpty(defaultStore.StoreID))
                            {
                                string filePath = string.Empty;
                                try { filePath = defaultStore.FilePath ?? string.Empty; } catch { }
                                bool isPst = false;
                                try { isPst = defaultStore.ExchangeStoreType == OutlookInterop.OlExchangeStoreType.olNotExchange; }
                                catch { isPst = filePath.EndsWith(".pst", StringComparison.OrdinalIgnoreCase); }

                                result.Add(new MailboxInfo
                                {
                                    StoreId = defaultStore.StoreID,
                                    DisplayName = defaultStore.DisplayName ?? "Default Store",
                                    FilePath = filePath,
                                    IsPst = isPst,
                                    IsSelected = true
                                });
                                Logger.Info($"[STORE DISCOVERY] Fallback: Discovered DefaultStore '{defaultStore.DisplayName}' (FilePath='{filePath}', IsPst={isPst})");
                            }
                        }
                        catch (Exception ex)
                        {
                            Logger.Warn($"[STORE DISCOVERY] Fallback DefaultStore check failed: {ex.Message}");
                        }
                        finally
                        {
                            ComReleaseHelper.Release(ref defaultStore);
                        }
                    }

                    // Fallback 2: Check top-level root folders in ns.Folders
                    if (result.Count == 0)
                    {
                        OutlookInterop.Folders? rootFolders = null;
                        try
                        {
                            rootFolders = ns.Folders;
                            int rootCount = rootFolders?.Count ?? 0;
                            for (int i = 1; i <= rootCount; i++)
                            {
                                OutlookInterop.MAPIFolder? rf = null;
                                try
                                {
                                    rf = rootFolders[i];
                                    if (rf != null && !string.IsNullOrEmpty(rf.StoreID) && !result.Any(r => string.Equals(r.StoreId, rf.StoreID, StringComparison.OrdinalIgnoreCase)))
                                    {
                                        result.Add(new MailboxInfo
                                        {
                                            StoreId = rf.StoreID,
                                            DisplayName = rf.Name ?? $"Store {i}",
                                            FilePath = string.Empty,
                                            IsPst = true,
                                            IsSelected = true
                                        });
                                        Logger.Info($"[STORE DISCOVERY] Fallback: Discovered top-level folder store '{rf.Name}' (StoreId='{rf.StoreID}')");
                                    }
                                }
                                catch { }
                                finally
                                {
                                    ComReleaseHelper.Release(ref rf);
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            Logger.Warn($"[STORE DISCOVERY] Fallback ns.Folders check failed: {ex.Message}");
                        }
                        finally
                        {
                            ComReleaseHelper.Release(ref rootFolders);
                        }
                    }

                    if (TestModeStoreFilter.IsTestModeActive)
                    {
                        result = TestModeStoreFilter.FilterMailboxes(result);
                    }
                    Logger.Info($"[STORE DISCOVERY] Completed. Total active stores discovered: {result.Count}");
                }
                catch (Exception ex)
                {
                    Logger.Error("Error retrieving Outlook mailboxes", ex);
                }
                finally
                {
                    ComReleaseHelper.Release(ref stores);
                    ComReleaseHelper.Release(ref ns);
                }

                return result;
            });
            }
            catch (ObjectDisposedException)
            {
                return new List<MailboxInfo>();
            }
        }

        public List<MailboxFolder> GetFolders(string storeId)
        {
            if (TestModeStoreFilter.IsTestModeActive && !TestModeStoreFilter.IsStoreIdApproved(storeId))
            {
                Logger.Warn($"[TEST MODE SAFEGUARD] GetFolders rejected unapproved store ID '{storeId}'.");
                return new List<MailboxFolder>();
            }
            return GetFoldersInternal(storeId);
        }

        private List<MailboxFolder> GetFoldersInternal(string storeId)
        {
            try
            {
                return _staThread.Invoke(() =>
            {
                var foldersList = new List<MailboxFolder>();
                var app = GetApplication();
                if (app == null) return foldersList;

                OutlookInterop.NameSpace? ns = null;
                OutlookInterop.Store? store = null;
                OutlookInterop.MAPIFolder? rootFolder = null;

                try
                {
                    ns = app.GetNamespace("MAPI");
                    foreach (OutlookInterop.Store s in ns.Stores)
                    {
                        if (string.Equals(s.StoreID, storeId, StringComparison.OrdinalIgnoreCase))
                        {
                            store = s;
                            break;
                        }
                        else
                        {
                            var tmp = s;
                            ComReleaseHelper.Release(ref tmp);
                        }
                    }

                    if (store != null)
                    {
                        rootFolder = store.GetRootFolder();
                        TraverseFolderTree(storeId, rootFolder, "", foldersList);
                    }
                    else
                    {
                        // Fallback: check top-level folders in ns.Folders
                        OutlookInterop.Folders? topFolders = null;
                        try
                        {
                            topFolders = ns.Folders;
                            int topCount = topFolders?.Count ?? 0;
                            for (int i = 1; i <= topCount; i++)
                            {
                                OutlookInterop.MAPIFolder? tf = null;
                                try
                                {
                                    tf = topFolders[i];
                                    if (tf != null && string.Equals(tf.StoreID, storeId, StringComparison.OrdinalIgnoreCase))
                                    {
                                        TraverseFolderTree(storeId, tf, "", foldersList);
                                        break;
                                    }
                                }
                                finally
                                {
                                    ComReleaseHelper.Release(ref tf);
                                }
                            }
                        }
                        finally
                        {
                            ComReleaseHelper.Release(ref topFolders);
                        }
                    }

                    Logger.Info($"[FOLDER DISCOVERY] Discovered {foldersList.Count} mail folder(s) for store '{storeId}'. Total items in folders: {foldersList.Sum(f => f.ItemCount):N0}");
                }
                catch (Exception ex)
                {
                    Logger.Error($"Error retrieving folders for store {storeId}", ex);
                }
                finally
                {
                    ComReleaseHelper.Release(ref rootFolder);
                    ComReleaseHelper.Release(ref store);
                    ComReleaseHelper.Release(ref ns);
                }

                return foldersList;
            });
            }
            catch (ObjectDisposedException)
            {
                return new List<MailboxFolder>();
            }
        }

        private void TraverseFolderTree(string storeId, OutlookInterop.MAPIFolder parentFolder, string parentId, List<MailboxFolder> list)
        {
            OutlookInterop.Folders? subFolders = null;
            try
            {
                OutlookInterop.OlItemType itemType = OutlookInterop.OlItemType.olMailItem;
                bool itemTypeOk = false;
                try
                {
                    itemType = parentFolder.DefaultItemType;
                    itemTypeOk = true;
                }
                catch (Exception ex)
                {
                    Logger.Warn($"Could not determine DefaultItemType for folder '{parentFolder.Name}': {ex.Message}");
                }

                if (itemTypeOk && itemType == OutlookInterop.OlItemType.olMailItem)
                {
                    int count = 0;
                    OutlookInterop.Items? items = null;
                    try
                    {
                        items = parentFolder.Items;
                        count = items?.Count ?? 0;
                    }
                    catch (Exception ex)
                    {
                        Logger.Warn($"Could not read Items.Count for folder '{parentFolder.Name}': {ex.Message}");
                    }
                    finally
                    {
                        ComReleaseHelper.Release(ref items);
                    }

                    // In PST files, the store root folder has parentId == "" and typically 0 items.
                    // Include folders that have a parentId or contain items.
                    if (!string.IsNullOrEmpty(parentId) || count > 0)
                    {
                        list.Add(new MailboxFolder
                        {
                            StoreId = storeId,
                            EntryId = parentFolder.EntryID,
                            Name = parentFolder.Name,
                            ParentId = parentId,
                            ItemCount = count
                        });
                        Logger.Info($"[FOLDER DISCOVERY] Mail folder: '{parentFolder.Name}' ({count:N0} items, EntryId='{parentFolder.EntryID}')");
                    }
                }
            }
            catch (Exception ex)
            {
                Logger.Warn($"Error evaluating folder '{parentFolder.Name}': {ex.Message}");
            }

            // Always traverse subfolders even if parent folder item check threw
            try
            {
                subFolders = parentFolder.Folders;
                int subFolderCount = subFolders?.Count ?? 0;
                for (int i = 1; i <= subFolderCount; i++)
                {
                    OutlookInterop.MAPIFolder? child = null;
                    try
                    {
                        child = subFolders[i];
                        TraverseFolderTree(storeId, child, parentFolder.EntryID, list);
                    }
                    catch (Exception ex)
                    {
                        Logger.Warn($"Failed reading subfolder {i} in '{parentFolder.Name}': {ex.Message}");
                    }
                    finally
                    {
                        ComReleaseHelper.Release(ref child);
                    }
                }
            }
            catch (Exception ex)
            {
                Logger.Warn($"Error traversing subfolders of '{parentFolder.Name}': {ex.Message}");
            }
            finally
            {
                ComReleaseHelper.Release(ref subFolders);
            }
        }

        private class FolderCursorSession : IDisposable
        {
            public string StoreId { get; set; } = string.Empty;
            public string FolderEntryId { get; set; } = string.Empty;
            public string FolderName { get; set; } = string.Empty;
            public OutlookInterop.NameSpace? NameSpace;
            public OutlookInterop.MAPIFolder? Folder;
            public OutlookInterop.Items? Items;
            public int TotalCount { get; set; }
            public int CurrentPosition { get; set; } // 1-based index of last item fetched/advanced
            public bool IsExhausted { get; set; }

            public void Dispose()
            {
                ComReleaseHelper.Release(ref Items);
                ComReleaseHelper.Release(ref Folder);
                ComReleaseHelper.Release(ref NameSpace);
            }
        }

        private FolderCursorSession? _activeCursorSession;

        public void ResetCursorSession()
        {
            try
            {
                _staThread.Invoke(() =>
                {
                    CloseActiveCursorSession();
                });
            }
            catch (ObjectDisposedException)
            {
                _activeCursorSession = null;
            }
        }

        private void CloseActiveCursorSession()
        {
            if (_activeCursorSession != null)
            {
                try
                {
                    _activeCursorSession.Dispose();
                }
                catch (Exception ex)
                {
                    Logger.Warn($"Error disposing cursor session: {ex.Message}");
                }
                finally
                {
                    _activeCursorSession = null;
                }
            }
        }

        public FolderBatchResult FetchFolderBatch(string storeId, string folderEntryId, int startIndex, int batchSize, Func<bool>? cancelCheck = null)
        {
            if (TestModeStoreFilter.IsTestModeActive && !TestModeStoreFilter.IsStoreIdApproved(storeId))
            {
                Logger.Warn($"[TEST MODE SAFEGUARD] FetchFolderBatch rejected unapproved store ID '{storeId}'.");
                return new FolderBatchResult
                {
                    StartIndex = startIndex,
                    LastProcessedIndex = startIndex - 1,
                    HasError = true,
                    ErrorMessage = $"Store '{storeId}' rejected by test mode safeguard."
                };
            }
            return FetchFolderBatchInternal(storeId, folderEntryId, startIndex, batchSize, cancelCheck);
        }

        private FolderBatchResult FetchFolderBatchInternal(string storeId, string folderEntryId, int startIndex, int batchSize, Func<bool>? cancelCheck = null)
        {
            try
            {
                return _staThread.Invoke(() =>
            {
                var result = new FolderBatchResult
                {
                    StartIndex = startIndex,
                    LastProcessedIndex = startIndex - 1,
                    ProcessedCount = 0,
                    IndexedCount = 0,
                    SkippedCount = 0
                };

                var app = GetApplication();
                if (app == null)
                {
                    result.HasError = true;
                    result.ErrorMessage = "Outlook application unavailable.";
                    return result;
                }

                try
                {
                    bool canReuseCursor = _activeCursorSession != null
                        && !_activeCursorSession.IsExhausted
                        && string.Equals(_activeCursorSession.StoreId, storeId, StringComparison.OrdinalIgnoreCase)
                        && string.Equals(_activeCursorSession.FolderEntryId, folderEntryId, StringComparison.OrdinalIgnoreCase)
                        && startIndex == _activeCursorSession.CurrentPosition + 1;

                    if (!canReuseCursor)
                    {
                        CloseActiveCursorSession();

                        OutlookInterop.NameSpace? ns = null;
                        OutlookInterop.MAPIFolder? folder = null;
                        OutlookInterop.Items? items = null;

                        try
                        {
                            ns = app.GetNamespace("MAPI");
                            folder = ns.GetFolderFromID(folderEntryId, storeId);
                            if (folder == null)
                            {
                                ComReleaseHelper.Release(ref ns);
                                result.HasError = true;
                                result.ErrorMessage = $"Folder not found: {folderEntryId}";
                                return result;
                            }

                            items = folder.Items;
                            int totalCount = items.Count;
                            string folderName = folder.Name;

                            if (totalCount == 0)
                            {
                                ComReleaseHelper.Release(ref items);
                                ComReleaseHelper.Release(ref folder);
                                ComReleaseHelper.Release(ref ns);
                                return result;
                            }

                            _activeCursorSession = new FolderCursorSession
                            {
                                StoreId = storeId,
                                FolderEntryId = folderEntryId,
                                FolderName = folderName,
                                NameSpace = ns,
                                Folder = folder,
                                Items = items,
                                TotalCount = totalCount,
                                CurrentPosition = 0
                            };
                        }
                        catch (Exception ex)
                        {
                            Logger.Error($"Error initializing cursor session for folder {folderEntryId}", ex);
                            ComReleaseHelper.Release(ref items);
                            ComReleaseHelper.Release(ref folder);
                            ComReleaseHelper.Release(ref ns);
                            result.HasError = true;
                            result.ErrorMessage = ex.Message;
                            throw;
                        }
                    }

                    var session = _activeCursorSession;
                    if (session == null || session.Items == null) return result;

                    int total = session.TotalCount;
                    if (startIndex > total)
                    {
                        CloseActiveCursorSession();
                        return result;
                    }

                    int endIdx = Math.Min(startIndex + batchSize - 1, total);

                    // If resuming at an offset beyond current cursor position, fast-forward to (startIndex - 1)
                    if (session.CurrentPosition < startIndex - 1)
                    {
                        object? skipItem = null;
                        int currentPos = session.CurrentPosition;

                        try
                        {
                            if (currentPos == 0)
                            {
                                skipItem = session.Items.GetFirst();
                                currentPos = 1;
                            }

                            while (currentPos < startIndex - 1 && skipItem != null)
                            {
                                if (cancelCheck != null && cancelCheck())
                                {
                                    ComReleaseHelper.Release(ref skipItem);
                                    session.CurrentPosition = currentPos;
                                    result.WasCancelled = true;
                                    result.LastProcessedIndex = startIndex - 1;
                                    return result;
                                }

                                ComReleaseHelper.Release(ref skipItem);
                                skipItem = session.Items.GetNext();
                                currentPos++;
                            }

                            ComReleaseHelper.Release(ref skipItem);
                            session.CurrentPosition = currentPos;
                        }
                        catch (Exception ex)
                        {
                            Logger.Error($"Error fast-forwarding cursor to offset {startIndex - 1}", ex);
                            ComReleaseHelper.Release(ref skipItem);
                            CloseActiveCursorSession();
                            result.HasError = true;
                            result.ErrorMessage = ex.Message;
                            throw;
                        }
                    }

                    // Read items from startIndex to endIdx sequentially
                    for (int i = startIndex; i <= endIdx; i++)
                    {
                        if (cancelCheck != null && cancelCheck())
                        {
                            result.WasCancelled = true;
                            break;
                        }

                        object? itemObj = null;
                        try
                        {
                            try
                            {
                                if (session.CurrentPosition == 0)
                                {
                                    itemObj = session.Items.GetFirst();
                                }
                                else
                                {
                                    itemObj = session.Items.GetNext();
                                }
                            }
                            catch (Exception comEx)
                            {
                                Logger.Warn($"Failed cursor advance at position {i} in folder {session.FolderName}: {comEx.Message}");
                                CloseActiveCursorSession();
                                if (i == startIndex)
                                {
                                    result.HasError = true;
                                    result.ErrorMessage = comEx.Message;
                                    throw;
                                }
                                else
                                {
                                    result.HasError = true;
                                    result.ErrorMessage = comEx.Message;
                                    break;
                                }
                            }

                            session.CurrentPosition = i;

                            if (itemObj == null)
                            {
                                session.IsExhausted = true;
                                break;
                            }

                            if (itemObj is OutlookInterop.MailItem mailItem)
                            {
                                try
                                {
                                    var doc = ExtractEmailDocument(mailItem, storeId, folderEntryId, session.FolderName);
                                    if (doc != null)
                                    {
                                        result.Items.Add(doc);
                                        result.IndexedCount++;
                                    }
                                    else
                                    {
                                        result.SkippedCount++;
                                    }
                                }
                                catch (Exception exExtract)
                                {
                                    Logger.Warn($"Failed extracting email {i} in folder {session.FolderName}: {exExtract.Message}");
                                    result.SkippedCount++;
                                }
                            }
                            else
                            {
                                result.SkippedCount++;
                            }

                            result.ProcessedCount++;
                            result.LastProcessedIndex = i;
                        }
                        finally
                        {
                            ComReleaseHelper.Release(ref itemObj);
                        }
                    }

                    if (session.CurrentPosition >= total || session.IsExhausted)
                    {
                        CloseActiveCursorSession();
                    }
                }
                catch (Exception ex)
                {
                    Logger.Error($"Error in FetchFolderBatch ({startIndex}..{startIndex + batchSize}) for folder {folderEntryId}", ex);
                    CloseActiveCursorSession();
                    if (!result.HasError)
                    {
                        result.HasError = true;
                        result.ErrorMessage = ex.Message;
                    }
                    throw;
                }

                return result;
            });
            }
            catch (ObjectDisposedException ex)
            {
                var result = FolderBatchResult.Empty(startIndex);
                result.HasError = true;
                result.ErrorMessage = $"OutlookAdapter has been disposed: {ex.Message}";
                return result;
            }
        }

        public List<EmailDocument> FetchFolderEmailsBatch(string storeId, string folderEntryId, int startIndex, int batchSize, Func<bool>? cancelCheck = null)
        {
            return FetchFolderBatch(storeId, folderEntryId, startIndex, batchSize, cancelCheck).Items;
        }

        public EmailDocument? FetchSingleEmail(string storeId, string entryId)
        {
            if (TestModeStoreFilter.IsTestModeActive && !TestModeStoreFilter.IsStoreIdApproved(storeId))
            {
                Logger.Warn($"[TEST MODE SAFEGUARD] FetchSingleEmail rejected unapproved store ID '{storeId}'.");
                return null;
            }
            return FetchSingleEmailInternal(storeId, entryId);
        }

        private EmailDocument? FetchSingleEmailInternal(string storeId, string entryId)
        {
            try
            {
                return _staThread.Invoke(() =>
                {
                    var sw = System.Diagnostics.Stopwatch.StartNew();
                    var app = GetApplication();
                    if (app == null) return null;

                    OutlookInterop.NameSpace? ns = null;
                    object? itemObj = null;

                    try
                    {
                        ns = app.GetNamespace("MAPI");
                        itemObj = ns.GetItemFromID(entryId, storeId);
                        if (itemObj is OutlookInterop.MailItem mailItem)
                        {
                            OutlookInterop.MAPIFolder? parent = mailItem.Parent as OutlookInterop.MAPIFolder;
                            string folderName = parent?.Name ?? "Inbox";
                            string folderEntryId = parent?.EntryID ?? "";
                            ComReleaseHelper.Release(ref parent);

                            // Requirement 3 & 9: Exclude Outbox item processing
                            if (IsOutboxFolder(storeId, folderEntryId) || folderName.Equals("Outbox", StringComparison.OrdinalIgnoreCase))
                            {
                                sw.Stop();
                                Logger.Info($"[COM] Operation=GetItemFromID EntryId={entryId} Folder=Outbox Action=SKIPPED_OUTBOX");
                                return null;
                            }

                            var doc = ExtractEmailDocument(mailItem, storeId, folderEntryId, folderName);
                            sw.Stop();
                            Logger.Info($"[COM] Operation=GetItemFromID Thread=FastMailSearch.STA DurationMs={sw.ElapsedMilliseconds}");
                            return doc;
                        }
                        
                        throw new InvalidOperationException("UNSUPPORTED_ITEM_TYPE");
                    }
                    catch (System.Runtime.InteropServices.COMException ex) when (ex.Message.IndexOf("could not be found", StringComparison.OrdinalIgnoreCase) >= 0 || ex.ErrorCode == -2147221233 || ex.ErrorCode == -2147221226)
                    {
                        Logger.Info($"[COM] Item permanently missing {entryId}: {ex.Message}");
                        throw new InvalidOperationException("ITEM_NOT_FOUND", ex);
                    }
                    catch (Exception ex)
                    {
                        Logger.Warn($"Failed to fetch single email {entryId}: {ex.Message}");
                        throw;
                    }
                    finally
                    {
                        ComReleaseHelper.Release(ref itemObj);
                        ComReleaseHelper.Release(ref ns);
                    }
                });
            }
            catch (ObjectDisposedException)
            {
                return null;
            }
        }

        public bool OpenEmailInOutlook(string storeId, string entryId)
        {
            if (TestModeStoreFilter.IsTestModeActive && !TestModeStoreFilter.IsStoreIdApproved(storeId))
            {
                Logger.Warn($"[TEST MODE SAFEGUARD] OpenEmailInOutlook rejected unapproved store ID '{storeId}'.");
                return false;
            }
            return OpenEmailInOutlookInternal(storeId, entryId);
        }

        private bool OpenEmailInOutlookInternal(string storeId, string entryId)
        {
            try
            {
                return _staThread.Invoke(() =>
                {
                    var app = GetApplication();
                    if (app == null) return false;

                    OutlookInterop.NameSpace? ns = null;
                    object? itemObj = null;

                    try
                    {
                        ns = app.GetNamespace("MAPI");
                        itemObj = ns.GetItemFromID(entryId, storeId);
                        if (itemObj is OutlookInterop.MailItem mailItem)
                        {
                            mailItem.Display(false); // Display non-modally
                            return true;
                        }
                    }
                    catch (COMException comEx)
                    {
                        Logger.Warn($"COM error opening item {entryId}: {comEx.Message}");
                        return false;
                    }
                    catch (Exception ex)
                    {
                        Logger.Error($"Error opening item {entryId} in Outlook", ex);
                        return false;
                    }
                    finally
                    {
                        ComReleaseHelper.Release(ref itemObj);
                        ComReleaseHelper.Release(ref ns);
                    }

                    return false;
                });
            }
            catch (ObjectDisposedException)
            {
                return false;
            }
        }

        public void SubscribeToEvents(Action<string, string, string> onItemAddedOrChanged, Action<string, string> onItemRemoved)
        {
            _staThread.Invoke(() =>
            {
                UnsubscribeFromEventsInternal();
                _onItemChangedCallback = onItemAddedOrChanged;
                _onItemRemovedCallback = onItemRemoved;

                var app = GetApplication();
                if (app == null) return;

                OutlookInterop.NameSpace? ns = null;
                try
                {
                    ns = app.GetNamespace("MAPI");
                    DiscoverDefaultFolders(ns, null);

                    foreach (OutlookInterop.Store store in ns.Stores)
                    {
                        try
                        {
                            DiscoverDefaultFolders(ns, store);
                            var root = store.GetRootFolder();
                            SubscribeFolderEvents(store.StoreID, root);
                            ComReleaseHelper.Release(ref root);
                        }
                        catch { }
                    }
                }
                catch (Exception ex)
                {
                    Logger.Error("Error subscribing to Outlook events", ex);
                }
                finally
                {
                    ComReleaseHelper.Release(ref ns);
                }
            });
        }

        private void SubscribeFolderEvents(string storeId, OutlookInterop.MAPIFolder folder)
        {
            OutlookInterop.Folders? subFolders = null;
            try
            {
                string folderEntryId = folder.EntryID;
                string folderName = folder.Name ?? string.Empty;

                // Requirement 3: Outbox Exclusion - Never subscribe to Outbox events
                if (IsOutboxFolder(storeId, folderEntryId) || folderName.Equals("Outbox", StringComparison.OrdinalIgnoreCase))
                {
                    Logger.Info($"[OUTLOOK EVENT] Type=FolderSubscribe Folder=Outbox EntryId={folderEntryId} Action=SKIPPED_OUTBOX_SUBSCRIPTION");
                    return;
                }

                if (folder.DefaultItemType == OutlookInterop.OlItemType.olMailItem)
                {
                    try
                    {
                        var items = folder.Items;
                        bool isSentItems = IsSentItemsFolderId(folderEntryId) || folderName.Equals("Sent Items", StringComparison.OrdinalIgnoreCase) || folderName.Equals("Sent", StringComparison.OrdinalIgnoreCase);

                        OutlookInterop.ItemsEvents_ItemAddEventHandler addHandler = (itemObj) =>
                        {
                            try
                            {
                                // Requirement 3 & 5: Non-blocking & Outbox guard
                                if (IsOutboxFolder(storeId, folderEntryId))
                                {
                                    Logger.Info($"[OUTLOOK EVENT] Type=ItemAdd Folder=Outbox Action=SKIPPED_OUTBOX");
                                    return;
                                }

                                if (itemObj is OutlookInterop.MailItem mailItem)
                                {
                                    string entryId = string.Empty;
                                    try
                                    {
                                        entryId = mailItem.EntryID ?? string.Empty;
                                    }
                                    catch { }

                                    if (!string.IsNullOrEmpty(entryId))
                                    {
                                        string folderLabel = isSentItems ? "Sent Items" : folderName;
                                        Logger.Info($"[OUTLOOK EVENT] Type=ItemAdd Folder={folderLabel} EntryId={entryId} Action=QUEUED_FOR_INDEX");
                                        _onItemChangedCallback?.Invoke(storeId, entryId, folderEntryId);
                                    }
                                }
                            }
                            catch (Exception ex)
                            {
                                Logger.Error("[OUTLOOK EVENT ERROR] Exception in ItemAdd handler: " + ex.Message, ex);
                            }
                        };

                        OutlookInterop.ItemsEvents_ItemChangeEventHandler changeHandler = (itemObj) =>
                        {
                            try
                            {
                                // Requirement 3 & 5: Non-blocking & Outbox guard
                                if (IsOutboxFolder(storeId, folderEntryId))
                                {
                                    Logger.Info($"[OUTLOOK EVENT] Type=ItemChange Folder=Outbox Action=SKIPPED_OUTBOX");
                                    return;
                                }

                                if (itemObj is OutlookInterop.MailItem mailItem)
                                {
                                    string entryId = string.Empty;
                                    try
                                    {
                                        entryId = mailItem.EntryID ?? string.Empty;
                                    }
                                    catch { }

                                    if (!string.IsNullOrEmpty(entryId))
                                    {
                                        string folderLabel = isSentItems ? "Sent Items" : folderName;
                                        Logger.Info($"[OUTLOOK EVENT] Type=ItemChange Folder={folderLabel} EntryId={entryId} Action=QUEUED_FOR_INDEX");
                                        _onItemChangedCallback?.Invoke(storeId, entryId, folderEntryId);
                                    }
                                }
                            }
                            catch (Exception ex)
                            {
                                Logger.Error("[OUTLOOK EVENT ERROR] Exception in ItemChange handler: " + ex.Message, ex);
                            }
                        };

                        if (SafeSubscribeItems(items, addHandler, changeHandler, out var sub) && sub != null)
                        {
                            _subscribedItemsList.Add(sub);
                        }
                    }
                    catch (Exception ex)
                    {
                        Logger.Warn($"[OUTLOOK EVENT] Folder event subscription failed for '{folderName}': {ex.Message}");
                    }
                }

                subFolders = folder.Folders;
                for (int i = 1; i <= subFolders.Count; i++)
                {
                    OutlookInterop.MAPIFolder? child = null;
                    try
                    {
                        child = subFolders[i];
                        SubscribeFolderEvents(storeId, child);
                    }
                    catch { }
                    finally
                    {
                        ComReleaseHelper.Release(ref child);
                    }
                }
            }
            catch { }
            finally
            {
                ComReleaseHelper.Release(ref subFolders);
            }
        }

        public void UnsubscribeFromEvents()
        {
            UnsubscribeFromEvents(TimeSpan.FromSeconds(3));
        }

        public void UnsubscribeFromEvents(TimeSpan timeout)
        {
            if (timeout <= TimeSpan.Zero)
            {
                UnsubscribeFromEventsInternal();
                return;
            }

            try
            {
                if (!_staThread.TryInvoke(() => UnsubscribeFromEventsInternal(), timeout))
                {
                    Logger.Warn("[OUTLOOK ADAPTER] UnsubscribeFromEvents timed out on STA thread; executing inline.");
                    UnsubscribeFromEventsInternal();
                }
            }
            catch (ObjectDisposedException)
            {
                UnsubscribeFromEventsInternal();
            }
        }

        private void UnsubscribeFromEventsInternal()
        {
            _onItemChangedCallback = null;
            _onItemRemovedCallback = null;
            foreach (var sub in _subscribedItemsList)
            {
                try
                {
                    sub.UnsubscribeAndRelease();
                }
                catch { }
            }
            _subscribedItemsList.Clear();
        }

        private EmailDocument? ExtractEmailDocument(OutlookInterop.MailItem mailItem, string storeId, string folderEntryId, string folderName)
        {
            if (mailItem == null) return null;

            string senderName = string.Empty;
            string senderEmail = string.Empty;
            string recipients = string.Empty;
            string cc = string.Empty;
            string bcc = string.Empty;
            string rawBody = string.Empty;

            if (TestModeStoreFilter.IsTestModeActive)
            {
                // In test mode out-of-process execution, avoid properties protected by Outlook Object Model Guard
                senderName = "Synthetic Test Sender";
                senderEmail = "synthetic@test.local";
                recipients = "recipient@test.local";
                rawBody = "Synthetic Test Email Body";
            }
            else
            {
                try { senderName = mailItem.SenderName ?? string.Empty; } catch { }
                try { senderEmail = mailItem.SenderEmailAddress ?? string.Empty; } catch { }
                try { recipients = mailItem.To ?? string.Empty; } catch { }
                try { cc = mailItem.CC ?? string.Empty; } catch { }
                try { bcc = mailItem.BCC ?? string.Empty; } catch { }
                try
                {
                    rawBody = mailItem.Body ?? string.Empty;
                    if (string.IsNullOrWhiteSpace(rawBody) && !string.IsNullOrWhiteSpace(mailItem.HTMLBody))
                    {
                        rawBody = StripHtml(mailItem.HTMLBody);
                    }
                }
                catch { }
            }

            var doc = new EmailDocument
            {
                StoreId = storeId,
                EntryId = mailItem.EntryID ?? string.Empty,
                ConversationId = mailItem.ConversationID ?? string.Empty,
                Subject = mailItem.Subject ?? string.Empty,
                SenderName = senderName,
                SenderEmail = senderEmail,
                Recipients = recipients,
                Cc = cc,
                Bcc = bcc,
                ReceivedDate = mailItem.ReceivedTime > DateTime.MinValue ? mailItem.ReceivedTime : DateTime.Now,
                SentDate = mailItem.SentOn > DateTime.MinValue ? mailItem.SentOn : DateTime.Now,
                FolderName = folderName,
                FolderEntryId = folderEntryId,
                HasAttachment = mailItem.Attachments.Count > 0,
                IsRead = !mailItem.UnRead,
                IsFlagged = mailItem.IsMarkedAsTask,
                Importance = (int)mailItem.Importance,
                Categories = mailItem.Categories ?? string.Empty,
                Size = mailItem.Size,
                LastModified = mailItem.LastModificationTime > DateTime.MinValue ? mailItem.LastModificationTime : DateTime.Now,
                IndexedDate = DateTime.Now
            };

            if (rawBody.Length > 1000000)
            {
                rawBody = rawBody.Substring(0, 1000000);
            }
            doc.BodyText = rawBody;

            if (doc.HasAttachment)
            {
                OutlookInterop.Attachments? atts = null;
                try
                {
                    atts = mailItem.Attachments;
                    for (int a = 1; a <= atts.Count; a++)
                    {
                        OutlookInterop.Attachment? att = null;
                        try
                        {
                            att = atts[a];
                            string name = att.FileName ?? att.DisplayName ?? $"Attachment_{a}";
                            string ext = Path.GetExtension(name);
                            doc.Attachments.Add(new EmailAttachment
                            {
                                Name = name,
                                Extension = ext,
                                Size = att.Size
                            });
                        }
                        catch { }
                        finally
                        {
                            ComReleaseHelper.Release(ref att);
                        }
                    }
                }
                catch { }
                finally
                {
                    ComReleaseHelper.Release(ref atts);
                }
            }

            return doc;
        }

        private string StripHtml(string html)
        {
            if (string.IsNullOrWhiteSpace(html)) return string.Empty;
            try
            {
                string step1 = Regex.Replace(html, @"<style[^>]*>[\s\S]*?</style>", "", RegexOptions.IgnoreCase);
                string step2 = Regex.Replace(step1, @"<script[^>]*>[\s\S]*?</script>", "", RegexOptions.IgnoreCase);
                string step3 = Regex.Replace(step2, @"<[^>]+>", " ", RegexOptions.IgnoreCase);
                return Regex.Replace(step3, @"\s+", " ").Trim();
            }
            catch
            {
                return html;
            }
        }

        private bool _disposed;
        public void Dispose()
        {
            Dispose(TimeSpan.FromSeconds(3));
        }

        public void Dispose(TimeSpan timeout)
        {
            if (_disposed) return;
            _disposed = true;
            ResetCursorSession();
            UnsubscribeFromEvents(timeout);
            int remainingMs = (int)Math.Max(0, timeout.TotalMilliseconds);
            _staThread.Dispose(remainingMs);
        }
    }
}
