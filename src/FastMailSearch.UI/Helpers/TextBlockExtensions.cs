using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Windows.Media;

namespace FastMailSearch.UI.Helpers
{
    public static class TextBlockExtensions
    {
        public static readonly DependencyProperty HighlightTextProperty =
            DependencyProperty.RegisterAttached(
                "HighlightText",
                typeof(string),
                typeof(TextBlockExtensions),
                new PropertyMetadata(null, OnHighlightChanged));

        public static readonly DependencyProperty QueryTextProperty =
            DependencyProperty.RegisterAttached(
                "QueryText",
                typeof(string),
                typeof(TextBlockExtensions),
                new PropertyMetadata(null, OnHighlightChanged));

        public static string GetHighlightText(DependencyObject obj)
        {
            return (string)obj.GetValue(HighlightTextProperty);
        }

        public static void SetHighlightText(DependencyObject obj, string value)
        {
            obj.SetValue(HighlightTextProperty, value);
        }

        public static string GetQueryText(DependencyObject obj)
        {
            return (string)obj.GetValue(QueryTextProperty);
        }

        public static void SetQueryText(DependencyObject obj, string value)
        {
            obj.SetValue(QueryTextProperty, value);
        }

        private static void OnHighlightChanged(DependencyObject d, DependencyPropertyChangedEventArgs e)
        {
            if (d is TextBlock textBlock)
            {
                string text = GetHighlightText(textBlock) ?? string.Empty;
                string rawQuery = GetQueryText(textBlock) ?? string.Empty;

                textBlock.Inlines.Clear();

                if (string.IsNullOrEmpty(text))
                {
                    return;
                }

                var highlightTerms = ExtractHighlightTerms(rawQuery);
                if (highlightTerms.Count == 0)
                {
                    textBlock.Inlines.Add(new Run(text));
                    return;
                }

                // Build a regex pattern for all terms
                var escapedTerms = new List<string>();
                foreach (var term in highlightTerms)
                {
                    escapedTerms.Add(Regex.Escape(term));
                }

                string pattern = "(" + string.Join("|", escapedTerms) + ")";
                var regex = new Regex(pattern, RegexOptions.IgnoreCase);

                string[] parts = regex.Split(text);
                var highlightBrush = new SolidColorBrush(Color.FromRgb(245, 158, 11)); // Amber highlight

                foreach (var part in parts)
                {
                    if (string.IsNullOrEmpty(part)) continue;

                    bool isMatch = false;
                    foreach (var term in highlightTerms)
                    {
                        if (part.Equals(term, StringComparison.OrdinalIgnoreCase))
                        {
                            isMatch = true;
                            break;
                        }
                    }

                    if (isMatch)
                    {
                        textBlock.Inlines.Add(new Run(part)
                        {
                            Background = highlightBrush,
                            Foreground = Brushes.Black,
                            FontWeight = FontWeights.Bold
                        });
                    }
                    else
                    {
                        textBlock.Inlines.Add(new Run(part));
                    }
                }
            }
        }

        private static List<string> ExtractHighlightTerms(string rawQuery)
        {
            var terms = new List<string>();
            if (string.IsNullOrWhiteSpace(rawQuery)) return terms;

            // Extract quoted phrases
            var quoteRegex = new Regex("\"([^\"]+)\"");
            var matches = quoteRegex.Matches(rawQuery);
            foreach (Match match in matches)
            {
                string phrase = match.Groups[1].Value.Trim();
                if (!string.IsNullOrEmpty(phrase)) terms.Add(phrase);
            }

            // Remove quoted phrases to parse remaining tokens
            string remaining = quoteRegex.Replace(rawQuery, " ");
            string[] tokens = remaining.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);

            foreach (var token in tokens)
            {
                if (token.Equals("OR", StringComparison.Ordinal) || token.Equals("AND", StringComparison.Ordinal))
                    continue;

                if (token.StartsWith("-")) continue;

                int colonIdx = token.IndexOf(':');
                if (colonIdx >= 0)
                {
                    string val = token.Substring(colonIdx + 1).Trim().Trim('"');
                    if (!string.IsNullOrEmpty(val))
                    {
                        terms.Add(val);
                    }
                }
                else
                {
                    string clean = token.Trim().Trim('"');
                    if (!string.IsNullOrEmpty(clean))
                    {
                        terms.Add(clean);
                    }
                }
            }

            return terms;
        }
    }
}
