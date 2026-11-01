using System;
using System.Collections.Generic;
using System.Text;
using FastMailSearch.Core.Models;

namespace FastMailSearch.Core.Search
{
    public static class SearchQueryParser
    {
        public static SearchQuery Parse(string rawQuery)
        {
            var searchQuery = new SearchQuery
            {
                RawQuery = rawQuery ?? string.Empty
            };

            if (string.IsNullOrWhiteSpace(rawQuery))
                return searchQuery;

            try
            {
                var (tokens, hasUnclosedQuote) = Tokenize(rawQuery!);
                if (hasUnclosedQuote)
                {
                    searchQuery.IsInvalidSyntax = true;
                    searchQuery.SyntaxErrorMessage = "Invalid search syntax: Unclosed quotation mark.";
                    return searchQuery;
                }

                var freeTextTokens = new List<string>();

                for (int i = 0; i < tokens.Count; i++)
                {
                    string token = tokens[i];

                    // Check for OR operator
                    if (token.Equals("OR", StringComparison.Ordinal))
                    {
                        searchQuery.BooleanOperator = "OR";
                        continue;
                    }

                    // Check for exclusion operator: -word
                    if (token.StartsWith("-") && token.Length > 1 && !token.Contains(":"))
                    {
                        string excludeTerm = token.Substring(1).Trim().Trim('"');
                        if (!string.IsNullOrEmpty(excludeTerm))
                        {
                            searchQuery.ExcludeTerms.Add(excludeTerm);
                        }
                        continue;
                    }

                    // Check for exact phrase in double quotes: "exact phrase"
                    if (token.StartsWith("\"") && token.EndsWith("\"") && token.Length >= 2)
                    {
                        string phrase = token.Substring(1, token.Length - 2).Trim();
                        if (!string.IsNullOrEmpty(phrase))
                        {
                            searchQuery.ExactPhrases.Add(phrase);
                            searchQuery.TextTerms.Add(phrase);
                            freeTextTokens.Add(token);
                        }
                        continue;
                    }

                    int colonIdx = token.IndexOf(':');
                    if (colonIdx >= 0)
                    {
                        string prefix = token.Substring(0, colonIdx).ToLowerInvariant().Trim();
                        string val = token.Substring(colonIdx + 1).Trim().Trim('"');

                        // Check for empty prefix value like "from:" or "subject:"
                        if (string.IsNullOrEmpty(val))
                        {
                            searchQuery.IsInvalidSyntax = true;
                            searchQuery.SyntaxErrorMessage = $"Invalid search syntax: Empty filter value for '{prefix}:'.";
                            return searchQuery;
                        }

                        switch (prefix)
                        {
                            case "from":
                                searchQuery.SenderFilter = val;
                                break;
                            case "to":
                                searchQuery.RecipientFilter = val;
                                break;
                            case "cc":
                                searchQuery.CcFilter = val;
                                break;
                            case "subject":
                                searchQuery.SubjectFilter = val;
                                break;
                            case "folder":
                                searchQuery.FolderFilter = val;
                                break;
                            case "store":
                            case "mailbox":
                                searchQuery.StoreFilter = val;
                                break;
                            case "filename":
                            case "file":
                            case "attachmentname":
                                searchQuery.FilenameFilter = val;
                                searchQuery.HasAttachmentFilter = true;
                                break;
                            case "has":
                                if (val.Equals("attachment", StringComparison.OrdinalIgnoreCase) || val.Equals("att", StringComparison.OrdinalIgnoreCase))
                                    searchQuery.HasAttachmentFilter = true;
                                break;
                            case "hasattachment":
                            case "attachment":
                                if (bool.TryParse(val, out bool hasAtt)) searchQuery.HasAttachmentFilter = hasAtt;
                                else if (val == "1" || val.Equals("yes", StringComparison.OrdinalIgnoreCase)) searchQuery.HasAttachmentFilter = true;
                                else if (val == "0" || val.Equals("no", StringComparison.OrdinalIgnoreCase)) searchQuery.HasAttachmentFilter = false;
                                break;
                            case "isread":
                            case "read":
                                if (bool.TryParse(val, out bool isRead)) searchQuery.IsReadFilter = isRead;
                                else if (val == "1" || val.Equals("yes", StringComparison.OrdinalIgnoreCase)) searchQuery.IsReadFilter = true;
                                else if (val == "0" || val.Equals("no", StringComparison.OrdinalIgnoreCase)) searchQuery.IsReadFilter = false;
                                break;
                            case "isflagged":
                            case "flagged":
                                if (bool.TryParse(val, out bool isFlagged)) searchQuery.IsFlaggedFilter = isFlagged;
                                else if (val == "1" || val.Equals("yes", StringComparison.OrdinalIgnoreCase)) searchQuery.IsFlaggedFilter = true;
                                else if (val == "0" || val.Equals("no", StringComparison.OrdinalIgnoreCase)) searchQuery.IsFlaggedFilter = false;
                                break;
                            case "after":
                            case "since":
                                var parsedAfter = DateParser.ParseSingleDate(val, isEnd: false);
                                if (!parsedAfter.HasValue)
                                {
                                    searchQuery.IsInvalidSyntax = true;
                                    searchQuery.SyntaxErrorMessage = $"Invalid search syntax: Unrecognized date format in '{prefix}:{val}'.";
                                    return searchQuery;
                                }
                                searchQuery.DateFrom = parsedAfter;
                                break;
                            case "before":
                            case "until":
                                var parsedBefore = DateParser.ParseSingleDate(val, isEnd: true);
                                if (!parsedBefore.HasValue)
                                {
                                    searchQuery.IsInvalidSyntax = true;
                                    searchQuery.SyntaxErrorMessage = $"Invalid search syntax: Unrecognized date format in '{prefix}:{val}'.";
                                    return searchQuery;
                                }
                                searchQuery.DateTo = parsedBefore;
                                break;
                            case "date":
                                var (df, dt) = DateParser.ParseRelativeDate(val);
                                if (df.HasValue) searchQuery.DateFrom = df;
                                if (dt.HasValue) searchQuery.DateTo = dt;
                                break;
                            default:
                                freeTextTokens.Add(token);
                                searchQuery.TextTerms.Add(token);
                                break;
                        }
                    }
                    else
                    {
                        freeTextTokens.Add(token);
                        searchQuery.TextTerms.Add(token);
                        searchQuery.IncludeTerms.Add(token);
                    }
                }

                searchQuery.FreeText = string.Join(" ", freeTextTokens).Trim();
            }
            catch
            {
                searchQuery.IsInvalidSyntax = true;
                searchQuery.SyntaxErrorMessage = "Invalid search syntax.";
            }

            return searchQuery;
        }

        private static (List<string> tokens, bool hasUnclosedQuote) Tokenize(string input)
        {
            var tokens = new List<string>();
            var sb = new StringBuilder();
            bool inQuotes = false;

            for (int i = 0; i < input.Length; i++)
            {
                char c = input[i];
                if (c == '"')
                {
                    inQuotes = !inQuotes;
                    sb.Append(c);
                }
                else if (char.IsWhiteSpace(c) && !inQuotes)
                {
                    if (sb.Length > 0)
                    {
                        tokens.Add(sb.ToString());
                        sb.Clear();
                    }
                }
                else
                {
                    sb.Append(c);
                }
            }

            if (sb.Length > 0)
            {
                tokens.Add(sb.ToString());
            }

            return (tokens, inQuotes);
        }
    }
}
