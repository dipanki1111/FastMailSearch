using System;
using System.Collections.Generic;
using FastMailSearch.Core.Models;

namespace FastMailSearch.Core.Search
{
    public static class ResultRanker
    {
        public static double CalculateRelevanceScore(SearchResultItem item, SearchQuery query)
        {
            double score = item.Rank; // Initial FTS rank (or 0)

            if (string.IsNullOrWhiteSpace(query.FreeText))
                return score;

            string text = query.FreeText.ToLowerInvariant().Trim();
            string[] terms = text.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);

            foreach (var term in terms)
            {
                if (term.Length <= 1) continue;

                // Match in Subject (10x boost)
                if (!string.IsNullOrEmpty(item.Subject) && item.Subject.ToLowerInvariant().Contains(term))
                {
                    score += 10.0;
                }

                // Match in SenderName or SenderEmail (5x boost)
                if ((!string.IsNullOrEmpty(item.SenderName) && item.SenderName.ToLowerInvariant().Contains(term)) ||
                    (!string.IsNullOrEmpty(item.SenderEmail) && item.SenderEmail.ToLowerInvariant().Contains(term)))
                {
                    score += 5.0;
                }

                // Match in Attachment names (3x boost)
                if (item.AttachmentNames != null)
                {
                    foreach (var attName in item.AttachmentNames)
                    {
                        if (attName.ToLowerInvariant().Contains(term))
                        {
                            score += 3.0;
                            break;
                        }
                    }
                }
            }

            // Recency boost (up to 2.0 points for recent emails within 30 days)
            if (item.ReceivedDate > DateTime.MinValue)
            {
                double daysOld = (DateTime.Now - item.ReceivedDate).TotalDays;
                if (daysOld >= 0 && daysOld < 30)
                {
                    score += 2.0 * (1.0 - (daysOld / 30.0));
                }
            }

            return score;
        }
    }
}
