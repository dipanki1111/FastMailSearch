using System;
using System.Collections.Generic;
using System.IO;
using System.Threading;
using FastMailSearch.Core.Models;
using Xunit;

namespace FastMailSearch.Core.Tests
{
    [Collection("Non-Parallel Environment Tests")]
    public class Phase36GSyntheticPstWhitelistTests
    {
        private static readonly SemaphoreSlim _envLock = new SemaphoreSlim(1, 1);

        #region Personal PST Permanent Rejection Tests

        [Fact]
        public void CanonicalPersonalPst_IsPermanentlyRejected_EvenIfWhitelistedByExactPath()
        {
            _envLock.Wait();
            string? origTest = Environment.GetEnvironmentVariable(TestModeStoreFilter.TestModeEnvironmentVariable);
            string? origAppr = Environment.GetEnvironmentVariable(TestModeStoreFilter.ApprovedStoresEnvironmentVariable);
            try
            {
                Environment.SetEnvironmentVariable(TestModeStoreFilter.TestModeEnvironmentVariable, "1");
                // Attacker/mistake tries to whitelist the exact canonical personal PST path
                Environment.SetEnvironmentVariable(TestModeStoreFilter.ApprovedStoresEnvironmentVariable, TestModeStoreFilter.HardcodedPersonalPstPath);
                TestModeStoreFilter.ResetTestMode();

                var personalMailbox = new MailboxInfo
                {
                    StoreId = "PERSONAL-STORE-ID",
                    DisplayName = "personal-test@example.com",
                    FilePath = TestModeStoreFilter.HardcodedPersonalPstPath,
                    IsPst = true
                };

                Assert.False(TestModeStoreFilter.IsStoreApproved(personalMailbox));
                Assert.False(TestModeStoreFilter.IsStoreIdApproved("PERSONAL-STORE-ID", "personal-test@example.com", TestModeStoreFilter.HardcodedPersonalPstPath, isPst: true));
            }
            finally
            {
                Environment.SetEnvironmentVariable(TestModeStoreFilter.TestModeEnvironmentVariable, origTest);
                Environment.SetEnvironmentVariable(TestModeStoreFilter.ApprovedStoresEnvironmentVariable, origAppr);
                TestModeStoreFilter.ResetTestMode();
                _envLock.Release();
            }
        }

        [Fact]
        public void CanonicalPersonalPst_IsPermanentlyRejected_WithRelativePathTraversal()
        {
            _envLock.Wait();
            string? origTest = Environment.GetEnvironmentVariable(TestModeStoreFilter.TestModeEnvironmentVariable);
            string? origAppr = Environment.GetEnvironmentVariable(TestModeStoreFilter.ApprovedStoresEnvironmentVariable);
            try
            {
                Environment.SetEnvironmentVariable(TestModeStoreFilter.TestModeEnvironmentVariable, "1");
                string traversalPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "Outlook Files", "..", "Outlook Files", "personal-test@example.com - test.pst");
                Environment.SetEnvironmentVariable(TestModeStoreFilter.ApprovedStoresEnvironmentVariable, traversalPath);
                TestModeStoreFilter.ResetTestMode();

                var traversalMailbox = new MailboxInfo
                {
                    StoreId = "TRAVERSAL-ID",
                    DisplayName = "My Mailbox",
                    FilePath = traversalPath,
                    IsPst = true
                };

                Assert.False(TestModeStoreFilter.IsStoreApproved(traversalMailbox));
                Assert.False(TestModeStoreFilter.IsPersonalPstPath(traversalPath) == false);
            }
            finally
            {
                Environment.SetEnvironmentVariable(TestModeStoreFilter.TestModeEnvironmentVariable, origTest);
                Environment.SetEnvironmentVariable(TestModeStoreFilter.ApprovedStoresEnvironmentVariable, origAppr);
                TestModeStoreFilter.ResetTestMode();
                _envLock.Release();
            }
        }

        [Fact]
        public void CanonicalPersonalPst_IsPermanentlyRejected_WithForwardSlashesAndMixedCase()
        {
            _envLock.Wait();
            string? origTest = Environment.GetEnvironmentVariable(TestModeStoreFilter.TestModeEnvironmentVariable);
            string? origAppr = Environment.GetEnvironmentVariable(TestModeStoreFilter.ApprovedStoresEnvironmentVariable);
            try
            {
                Environment.SetEnvironmentVariable(TestModeStoreFilter.TestModeEnvironmentVariable, "1");
                string mixedPath = (Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "outlook files", "PERSONAL-TEST@EXAMPLE.COM - TEST.PST") + " ").Replace('\\', '/');
                Environment.SetEnvironmentVariable(TestModeStoreFilter.ApprovedStoresEnvironmentVariable, mixedPath);
                TestModeStoreFilter.ResetTestMode();

                var mixedMailbox = new MailboxInfo
                {
                    StoreId = "MIXED-ID",
                    DisplayName = "personal-test@example.com",
                    FilePath = mixedPath,
                    IsPst = true
                };

                Assert.False(TestModeStoreFilter.IsStoreApproved(mixedMailbox));
            }
            finally
            {
                Environment.SetEnvironmentVariable(TestModeStoreFilter.TestModeEnvironmentVariable, origTest);
                Environment.SetEnvironmentVariable(TestModeStoreFilter.ApprovedStoresEnvironmentVariable, origAppr);
                TestModeStoreFilter.ResetTestMode();
                _envLock.Release();
            }
        }

        [Fact]
        public void CanonicalPersonalPst_IsPermanentlyRejected_ByPersonalDisplayName()
        {
            _envLock.Wait();
            try
            {
                TestModeStoreFilter.SetTestMode(true, new[] { "personal-test@example.com" });

                var personalNameMailbox = new MailboxInfo
                {
                    StoreId = "SOME-ID-99",
                    DisplayName = "personal-test@example.com",
                    FilePath = @"D:\SomeOtherLocation\mail.pst",
                    IsPst = true
                };

                Assert.False(TestModeStoreFilter.IsStoreApproved(personalNameMailbox));
            }
            finally
            {
                TestModeStoreFilter.ResetTestMode();
                _envLock.Release();
            }
        }

        [Fact]
        public void CanonicalPersonalPst_IsPermanentlyRejected_ByPersonalStoreId()
        {
            _envLock.Wait();
            try
            {
                TestModeStoreFilter.SetTestMode(true, new[] { TestModeStoreFilter.CanonicalPersonalStoreId });

                var personalIdMailbox = new MailboxInfo
                {
                    StoreId = TestModeStoreFilter.CanonicalPersonalStoreId,
                    DisplayName = "Personal Account",
                    FilePath = @"D:\SomeOtherLocation\mail.pst",
                    IsPst = true
                };

                Assert.False(TestModeStoreFilter.IsStoreApproved(personalIdMailbox));
            }
            finally
            {
                TestModeStoreFilter.ResetTestMode();
                _envLock.Release();
            }
        }

        #endregion

        #region OST Rejection Tests

        [Fact]
        public void AllOstFiles_ArePermanentlyRejected_EvenIfWhitelistedByPathOrId()
        {
            _envLock.Wait();
            try
            {
                string ostPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "Microsoft", "Outlook", "test_exchange.ost");
                TestModeStoreFilter.SetTestMode(true, new[] { ostPath, "OST-STORE-101" });

                var ostMailbox = new MailboxInfo
                {
                    StoreId = "OST-STORE-101",
                    DisplayName = "Exchange Account",
                    FilePath = ostPath,
                    IsPst = false
                };

                Assert.False(TestModeStoreFilter.IsStoreApproved(ostMailbox));
                Assert.False(TestModeStoreFilter.IsStoreIdApproved("OST-STORE-101", "Exchange Account", ostPath, isPst: false));
            }
            finally
            {
                TestModeStoreFilter.ResetTestMode();
                _envLock.Release();
            }
        }

        #endregion

        #region Synthetic PST Whitelist & Unapproved PST Rejection Tests

        [Fact]
        public void ApprovedSyntheticPst_IsAllowed_ByExactPath()
        {
            _envLock.Wait();
            string? origTest = Environment.GetEnvironmentVariable(TestModeStoreFilter.TestModeEnvironmentVariable);
            string? origAppr = Environment.GetEnvironmentVariable(TestModeStoreFilter.ApprovedStoresEnvironmentVariable);
            try
            {
                string syntheticPst = Path.Combine(Path.GetTempPath(), "FastMailSearch_Test", "SyntheticTest.pst");
                Environment.SetEnvironmentVariable(TestModeStoreFilter.TestModeEnvironmentVariable, "1");
                Environment.SetEnvironmentVariable(TestModeStoreFilter.ApprovedStoresEnvironmentVariable, syntheticPst);
                TestModeStoreFilter.ResetTestMode();

                var syntheticMailbox = new MailboxInfo
                {
                    StoreId = "SYNTHETIC-STORE-001",
                    DisplayName = "Synthetic Integration Mailbox",
                    FilePath = syntheticPst,
                    IsPst = true
                };

                Assert.True(TestModeStoreFilter.IsStoreApproved(syntheticMailbox));
                Assert.True(TestModeStoreFilter.IsStoreIdApproved("SYNTHETIC-STORE-001"));
            }
            finally
            {
                Environment.SetEnvironmentVariable(TestModeStoreFilter.TestModeEnvironmentVariable, origTest);
                Environment.SetEnvironmentVariable(TestModeStoreFilter.ApprovedStoresEnvironmentVariable, origAppr);
                TestModeStoreFilter.ResetTestMode();
                _envLock.Release();
            }
        }

        [Fact]
        public void ApprovedSyntheticPst_IsAllowed_WithNormalizedPathVariations()
        {
            _envLock.Wait();
            try
            {
                string basePst = Path.Combine(Path.GetTempPath(), "FastMailSearch_Test", "SyntheticTest.pst");
                TestModeStoreFilter.SetTestMode(true, new[] { basePst });

                string variationPath = (Path.Combine(Path.GetTempPath(), "fastmailsearch_test", "..", "FastMailSearch_Test", "SyntheticTest.pst") + " ").Replace('\\', '/');

                var syntheticMailbox = new MailboxInfo
                {
                    StoreId = "SYNTHETIC-VAR-001",
                    DisplayName = "Synthetic Variation",
                    FilePath = variationPath,
                    IsPst = true
                };

                Assert.True(TestModeStoreFilter.IsStoreApproved(syntheticMailbox));
            }
            finally
            {
                TestModeStoreFilter.ResetTestMode();
                _envLock.Release();
            }
        }

        [Fact]
        public void ApprovedSyntheticPst_IsAllowed_ByStoreId()
        {
            _envLock.Wait();
            try
            {
                string storeId = "SYNTHETIC-STORE-ID-9999";
                TestModeStoreFilter.SetTestMode(true, new[] { storeId });

                var syntheticMailbox = new MailboxInfo
                {
                    StoreId = storeId,
                    DisplayName = "Dedicated Synthetic Test Store",
                    FilePath = Path.Combine(Path.GetTempPath(), "DedicatedSynthetic.pst"),
                    IsPst = true
                };

                Assert.True(TestModeStoreFilter.IsStoreApproved(syntheticMailbox));
                Assert.True(TestModeStoreFilter.IsStoreIdApproved(storeId));
            }
            finally
            {
                TestModeStoreFilter.ResetTestMode();
                _envLock.Release();
            }
        }

        [Fact]
        public void UnapprovedPst_IsRejected_WhenAnotherSyntheticPstIsApproved()
        {
            _envLock.Wait();
            try
            {
                string approvedPst = Path.Combine(Path.GetTempPath(), "FastMailSearch_Test", "ApprovedSynthetic.pst");
                TestModeStoreFilter.SetTestMode(true, new[] { approvedPst });

                var unapprovedMailbox = new MailboxInfo
                {
                    StoreId = "UNAPPROVED-PST-002",
                    DisplayName = "Random PST File",
                    FilePath = Path.Combine(Path.GetTempPath(), "FastMailSearch_Test", "UnapprovedOther.pst"),
                    IsPst = true
                };

                Assert.False(TestModeStoreFilter.IsStoreApproved(unapprovedMailbox));
                Assert.False(TestModeStoreFilter.IsStoreIdApproved("UNAPPROVED-PST-002"));
            }
            finally
            {
                TestModeStoreFilter.ResetTestMode();
                _envLock.Release();
            }
        }

        [Fact]
        public void EmptyApprovedList_FailsClosed_RejectsSyntheticPst()
        {
            _envLock.Wait();
            string? origTest = Environment.GetEnvironmentVariable(TestModeStoreFilter.TestModeEnvironmentVariable);
            string? origAppr = Environment.GetEnvironmentVariable(TestModeStoreFilter.ApprovedStoresEnvironmentVariable);
            try
            {
                Environment.SetEnvironmentVariable(TestModeStoreFilter.TestModeEnvironmentVariable, "1");
                Environment.SetEnvironmentVariable(TestModeStoreFilter.ApprovedStoresEnvironmentVariable, "");
                TestModeStoreFilter.ResetTestMode();

                var syntheticMailbox = new MailboxInfo
                {
                    StoreId = "SYNTHETIC-EMPTY-001",
                    DisplayName = "Synthetic Store",
                    FilePath = Path.Combine(Path.GetTempPath(), "FastMailSearch_Test", "Synthetic.pst"),
                    IsPst = true
                };

                Assert.False(TestModeStoreFilter.IsStoreApproved(syntheticMailbox));
                Assert.False(TestModeStoreFilter.IsStoreIdApproved("SYNTHETIC-EMPTY-001"));
            }
            finally
            {
                Environment.SetEnvironmentVariable(TestModeStoreFilter.TestModeEnvironmentVariable, origTest);
                Environment.SetEnvironmentVariable(TestModeStoreFilter.ApprovedStoresEnvironmentVariable, origAppr);
                TestModeStoreFilter.ResetTestMode();
                _envLock.Release();
            }
        }

        #endregion

        #region Production Behavior & Path Normalization Tests

        [Fact]
        public void TestModeInactive_AllowsAllStores_PreservingProduction()
        {
            _envLock.Wait();
            string? origTest = Environment.GetEnvironmentVariable(TestModeStoreFilter.TestModeEnvironmentVariable);
            try
            {
                Environment.SetEnvironmentVariable(TestModeStoreFilter.TestModeEnvironmentVariable, null);
                TestModeStoreFilter.ResetTestMode();

                Assert.False(TestModeStoreFilter.IsTestModeActive);

                var personalMailbox = new MailboxInfo
                {
                    StoreId = TestModeStoreFilter.CanonicalPersonalStoreId,
                    DisplayName = "personal-test@example.com",
                    FilePath = TestModeStoreFilter.HardcodedPersonalPstPath,
                    IsPst = true
                };

                var ostMailbox = new MailboxInfo
                {
                    StoreId = "OST-001",
                    DisplayName = "Exchange Account",
                    FilePath = @"C:\Mail\work.ost",
                    IsPst = false
                };

                // Production mode: all stores permitted
                Assert.True(TestModeStoreFilter.IsStoreApproved(personalMailbox));
                Assert.True(TestModeStoreFilter.IsStoreApproved(ostMailbox));

                var filtered = TestModeStoreFilter.FilterMailboxes(new[] { personalMailbox, ostMailbox });
                Assert.Equal(2, filtered.Count);
            }
            finally
            {
                Environment.SetEnvironmentVariable(TestModeStoreFilter.TestModeEnvironmentVariable, origTest);
                TestModeStoreFilter.ResetTestMode();
                _envLock.Release();
            }
        }

        [Fact]
        public void PathNormalization_RejectsPartialOrAmbiguousSubstringMatches()
        {
            _envLock.Wait();
            try
            {
                // Approved path is a specific synthetic PST
                string approvedPst = Path.Combine(Path.GetTempPath(), "SyntheticTest.pst");
                TestModeStoreFilter.SetTestMode(true, new[] { approvedPst });

                // Similar prefix or suffix or substring must NOT match
                var prefixMatch = new MailboxInfo
                {
                    StoreId = "P1",
                    DisplayName = "Fake Prefix",
                    FilePath = Path.Combine(Path.GetTempPath(), "SyntheticTest.pst.bak"),
                    IsPst = true
                };

                var subfolderMatch = new MailboxInfo
                {
                    StoreId = "P2",
                    DisplayName = "Fake Subfolder",
                    FilePath = Path.Combine(Path.GetTempPath(), "Sub", "SyntheticTest.pst"),
                    IsPst = true
                };

                var shortMatch = new MailboxInfo
                {
                    StoreId = "P3",
                    DisplayName = "Fake Short",
                    FilePath = Path.Combine(Path.GetTempPath(), "Test.pst"),
                    IsPst = true
                };

                Assert.False(TestModeStoreFilter.IsStoreApproved(prefixMatch));
                Assert.False(TestModeStoreFilter.IsStoreApproved(subfolderMatch));
                Assert.False(TestModeStoreFilter.IsStoreApproved(shortMatch));
            }
            finally
            {
                TestModeStoreFilter.ResetTestMode();
                _envLock.Release();
            }
        }

        [Fact]
        public void FilterMailboxes_KeepsOnlyApprovedSyntheticPst_DropsPersonalPstAndOst()
        {
            _envLock.Wait();
            try
            {
                string approvedSyntheticPst = Path.Combine(Path.GetTempPath(), "FastMailSearch_Test", "SyntheticTest.pst");
                TestModeStoreFilter.SetTestMode(true, new[] { approvedSyntheticPst });

                var personalPst = new MailboxInfo
                {
                    StoreId = "PERSONAL-001",
                    DisplayName = "personal-test@example.com",
                    FilePath = TestModeStoreFilter.HardcodedPersonalPstPath,
                    IsPst = true
                };

                var ostMailbox = new MailboxInfo
                {
                    StoreId = "OST-001",
                    DisplayName = "Exchange Live",
                    FilePath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "Outlook.ost"),
                    IsPst = false
                };

                var unapprovedPst = new MailboxInfo
                {
                    StoreId = "UNAPPROVED-001",
                    DisplayName = "Other Archive",
                    FilePath = @"D:\Archives\Archive2025.pst",
                    IsPst = true
                };

                var approvedSynthetic = new MailboxInfo
                {
                    StoreId = "SYNTHETIC-APPROVED-001",
                    DisplayName = "Dedicated Synthetic Test Store",
                    FilePath = approvedSyntheticPst,
                    IsPst = true
                };

                var inputList = new List<MailboxInfo> { personalPst, ostMailbox, unapprovedPst, approvedSynthetic };
                var filtered = TestModeStoreFilter.FilterMailboxes(inputList);

                Assert.Single(filtered);
                Assert.Same(approvedSynthetic, filtered[0]);
            }
            finally
            {
                TestModeStoreFilter.ResetTestMode();
                _envLock.Release();
            }
        }

        #endregion
    }
}

