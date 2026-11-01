using System;
using FastMailSearch.Core.Search;
using Xunit;

namespace FastMailSearch.Core.Tests
{
    public class SearchQueryParserTests
    {
        [Fact]
        public void Parse_SimpleFreeText_ReturnsFreeText()
        {
            var q = SearchQueryParser.Parse("invoice ABC");
            Assert.Equal("invoice ABC", q.FreeText);
        }

        [Fact]
        public void Parse_QuotedString_PreservesQuotesInFreeText()
        {
            var q = SearchQueryParser.Parse("\"invoice 2026\"");
            Assert.Equal("\"invoice 2026\"", q.FreeText);
        }

        [Fact]
        public void Parse_FieldPrefixes_ExtractsTypedFilters()
        {
            var q = SearchQueryParser.Parse("from:abc.com subject:invoice hasattachment:true isread:false after:2026-01-01 payment");

            Assert.Equal("abc.com", q.SenderFilter);
            Assert.Equal("invoice", q.SubjectFilter);
            Assert.True(q.HasAttachmentFilter);
            Assert.False(q.IsReadFilter);
            Assert.NotNull(q.DateFrom);
            Assert.Equal(new DateTime(2026, 1, 1), q.DateFrom.Value.Date);
            Assert.Equal("payment", q.FreeText);
        }

        [Fact]
        public void Parse_RelativeDates_ParsesTodayAndYesterday()
        {
            var qToday = SearchQueryParser.Parse("date:today");
            Assert.NotNull(qToday.DateFrom);
            Assert.Equal(DateTime.Today, qToday.DateFrom.Value.Date);

            var qYesterday = SearchQueryParser.Parse("date:yesterday");
            Assert.NotNull(qYesterday.DateFrom);
            Assert.Equal(DateTime.Today.AddDays(-1), qYesterday.DateFrom.Value.Date);
        }

        [Fact]
        public void Parse_AdvancedSearchOperators_ExtractsAllFilters()
        {
            var q = SearchQueryParser.Parse("from:kfintech to:abc@gmail.com subject:IPO after:2026-08-01 before:2026-08-15 attachment:yes folder:Inbox");

            Assert.Equal("kfintech", q.SenderFilter);
            Assert.Equal("abc@gmail.com", q.RecipientFilter);
            Assert.Equal("IPO", q.SubjectFilter);
            Assert.NotNull(q.DateFrom);
            Assert.Equal(new DateTime(2026, 8, 1), q.DateFrom.Value.Date);
            Assert.NotNull(q.DateTo);
            Assert.Equal(new DateTime(2026, 8, 15), q.DateTo.Value.Date);
            Assert.True(q.HasAttachmentFilter);
            Assert.Equal("Inbox", q.FolderFilter);
        }

        [Fact]
        public void Parse_ExactPhrase_ExtractsExactPhrases()
        {
            var q = SearchQueryParser.Parse("from:john@example.com \"project meeting\"");
            Assert.Equal("john@example.com", q.SenderFilter);
            Assert.Contains("project meeting", q.ExactPhrases);
        }

        [Fact]
        public void Parse_CC_And_Filename_ExtractsFilters()
        {
            var q = SearchQueryParser.Parse("cc:manager@test.com filename:report.pdf has:attachment");
            Assert.Equal("manager@test.com", q.CcFilter);
            Assert.Equal("report.pdf", q.FilenameFilter);
            Assert.True(q.HasAttachmentFilter);
        }

        [Fact]
        public void Parse_ExclusionAndBoolean_ParsesExcludeTermsAndOperator()
        {
            var q = SearchQueryParser.Parse("invoice OR quotation -spam");
            Assert.Equal("OR", q.BooleanOperator);
            Assert.Contains("spam", q.ExcludeTerms);
        }

        [Fact]
        public void Parse_InvalidSyntax_SetsIsInvalidSyntaxFlag()
        {
            var qUnclosed = SearchQueryParser.Parse("\"unclosed phrase");
            Assert.True(qUnclosed.IsInvalidSyntax);

            var qEmptyFrom = SearchQueryParser.Parse("from:");
            Assert.True(qEmptyFrom.IsInvalidSyntax);

            var qEmptySubject = SearchQueryParser.Parse("subject:");
            Assert.True(qEmptySubject.IsInvalidSyntax);

            var qEmptyAfter = SearchQueryParser.Parse("after:");
            Assert.True(qEmptyAfter.IsInvalidSyntax);

            var qEmptyBefore = SearchQueryParser.Parse("before:");
            Assert.True(qEmptyBefore.IsInvalidSyntax);

            var qEmptyFilename = SearchQueryParser.Parse("filename:");
            Assert.True(qEmptyFilename.IsInvalidSyntax);
        }

        [Fact]
        public void Parse_Phase1_All15SyntaxItems_Supported()
        {
            // 1. Normal text
            var q1 = SearchQueryParser.Parse("kfintech");
            Assert.Equal("kfintech", q1.FreeText);

            // 2. Exact phrase
            var q2 = SearchQueryParser.Parse("\"project meeting\"");
            Assert.Contains("project meeting", q2.ExactPhrases);

            // 3. From
            var q3 = SearchQueryParser.Parse("from:john@example.com");
            Assert.Equal("john@example.com", q3.From);

            // 4. To
            var q4 = SearchQueryParser.Parse("to:accounts@example.com");
            Assert.Equal("accounts@example.com", q4.To);

            // 5. CC
            var q5 = SearchQueryParser.Parse("cc:manager@example.com");
            Assert.Equal("manager@example.com", q5.Cc);

            // 6. Subject
            var q6 = SearchQueryParser.Parse("subject:invoice");
            Assert.Equal("invoice", q6.Subject);

            // 7. Attachment
            var q7 = SearchQueryParser.Parse("has:attachment");
            Assert.True(q7.HasAttachment);

            // 8. Filename
            var q8 = SearchQueryParser.Parse("filename:report.pdf");
            Assert.Equal("report.pdf", q8.Filename);
            Assert.True(q8.HasAttachment);

            // 9. Folder
            var q9 = SearchQueryParser.Parse("folder:inbox");
            Assert.Equal("inbox", q9.Folder);

            // 10. Date (after / before)
            var q10a = SearchQueryParser.Parse("after:2026-01-01");
            Assert.Equal(new DateTime(2026, 1, 1), q10a.AfterDate.Value.Date);

            var q10b = SearchQueryParser.Parse("before:2026-08-15");
            Assert.Equal(new DateTime(2026, 8, 15), q10b.BeforeDate.Value.Date);

            // 11. Date range
            var q11 = SearchQueryParser.Parse("after:2026-01-01 before:2026-03-31");
            Assert.Equal(new DateTime(2026, 1, 1), q11.AfterDate.Value.Date);
            Assert.Equal(new DateTime(2026, 3, 31), q11.BeforeDate.Value.Date);

            // 12. Boolean
            var q12 = SearchQueryParser.Parse("invoice OR quotation");
            Assert.Equal("OR", q12.BooleanOperator);
            Assert.Contains("invoice", q12.IncludeTerms);
            Assert.Contains("quotation", q12.IncludeTerms);

            // 13. Exclusion
            var q13 = SearchQueryParser.Parse("invoice -spam");
            Assert.Contains("invoice", q13.IncludeTerms);
            Assert.Contains("spam", q13.ExcludeTerms);

            // 14. Combined
            var q14 = SearchQueryParser.Parse("from:john@example.com subject:invoice after:2026-01-01");
            Assert.Equal("john@example.com", q14.From);
            Assert.Equal("invoice", q14.Subject);
            Assert.Equal(new DateTime(2026, 1, 1), q14.AfterDate.Value.Date);

            // 15. Exact phrase + filter
            var q15 = SearchQueryParser.Parse("from:john@example.com \"payment received\"");
            Assert.Equal("john@example.com", q15.From);
            Assert.Contains("payment received", q15.ExactPhrases);
        }
    }
}
