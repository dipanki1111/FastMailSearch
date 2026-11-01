using System;
using System.Threading;
using System.Threading.Tasks;
using FastMailSearch.Core.Interfaces;
using FastMailSearch.Core.Models;
using FastMailSearch.Core.Services;
using Xunit;

namespace FastMailSearch.Core.Tests
{
    public class SearchServiceLicensingEnforcementTests
    {
        [Fact]
        public async Task ExecuteSearchAsync_AllowsSearch_When_LicenseStatusIsValid()
        {
            var mockLicensing = new MockLicensingService(LicenseStatus.Valid);
            var service = new SearchService(new MockSearchEngine(), new MockRepository(), mockLicensing);

            var result = await service.ExecuteSearchAsync("test query");

            Assert.NotNull(result);
            Assert.Equal(100, result.TotalCount);
            Assert.Equal(1, mockLicensing.GetStatusCallCount);
        }

        [Fact]
        public async Task ExecuteSearchAsync_AllowsSearch_When_LicenseStatusIsTrial()
        {
            var mockLicensing = new MockLicensingService(LicenseStatus.Trial);
            var service = new SearchService(new MockSearchEngine(), new MockRepository(), mockLicensing);

            var result = await service.ExecuteSearchAsync("test query");

            Assert.NotNull(result);
            Assert.Equal(100, result.TotalCount);
            Assert.Equal(1, mockLicensing.GetStatusCallCount);
        }

        [Fact]
        public async Task ExecuteSearchAsync_RejectsSearch_When_LicenseStatusIsExpired()
        {
            var mockLicensing = new MockLicensingService(LicenseStatus.Expired);
            var service = new SearchService(new MockSearchEngine(), new MockRepository(), mockLicensing);

            var ex = await Assert.ThrowsAsync<InvalidOperationException>(() => service.ExecuteSearchAsync("test query"));

            Assert.Contains("expired", ex.Message, StringComparison.OrdinalIgnoreCase);
            Assert.Equal(1, mockLicensing.GetStatusCallCount);
        }

        [Fact]
        public async Task ExecuteSearchAsync_RejectsSearch_When_LicenseStatusIsInvalidKey()
        {
            var mockLicensing = new MockLicensingService(LicenseStatus.InvalidKey);
            var service = new SearchService(new MockSearchEngine(), new MockRepository(), mockLicensing);

            var ex = await Assert.ThrowsAsync<InvalidOperationException>(() => service.ExecuteSearchAsync("test query"));

            Assert.Contains("expired", ex.Message, StringComparison.OrdinalIgnoreCase);
            Assert.Equal(1, mockLicensing.GetStatusCallCount);
        }

        [Fact]
        public async Task ExecuteSearchAsync_RejectsSearch_When_LicenseStatusIsNonEntitledOrUnknown()
        {
            var mockLicensing = new MockLicensingService((LicenseStatus)999);
            var service = new SearchService(new MockSearchEngine(), new MockRepository(), mockLicensing);

            var ex = await Assert.ThrowsAsync<InvalidOperationException>(() => service.ExecuteSearchAsync("test query"));

            Assert.Contains("expired", ex.Message, StringComparison.OrdinalIgnoreCase);
            Assert.Equal(1, mockLicensing.GetStatusCallCount);
        }

        [Fact]
        public async Task ExecuteSearchAsync_RejectsSearch_When_ConstructedWithoutLicensingService_MissingDependency()
        {
#pragma warning disable CS0618 // Type or member is obsolete
            var service = new SearchService(new MockSearchEngine(), new MockRepository());
#pragma warning restore CS0618

            var ex = await Assert.ThrowsAsync<InvalidOperationException>(() => service.ExecuteSearchAsync("test query"));

            Assert.Contains("licensing", ex.Message, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("expired", ex.Message, StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public void SearchService_Constructor_ThrowsArgumentNullException_WhenLicensingServiceIsNull()
        {
            var ex = Assert.Throws<ArgumentNullException>(() =>
                new SearchService(new MockSearchEngine(), new MockRepository(), null!));

            Assert.Equal("licensingService", ex.ParamName);
        }

        [Fact]
        public void SearchService_Constructor_ThrowsArgumentNullException_WhenSearchEngineIsNull()
        {
            var mockLicensing = new MockLicensingService(LicenseStatus.Valid);
            var ex = Assert.Throws<ArgumentNullException>(() =>
                new SearchService(null!, new MockRepository(), mockLicensing));

            Assert.Equal("searchEngine", ex.ParamName);
        }

        [Fact]
        public void SearchService_Constructor_ThrowsArgumentNullException_WhenRepositoryIsNull()
        {
            var mockLicensing = new MockLicensingService(LicenseStatus.Valid);
            var ex = Assert.Throws<ArgumentNullException>(() =>
                new SearchService(new MockSearchEngine(), null!, mockLicensing));

            Assert.Equal("repository", ex.ParamName);
        }

        [Fact]
        public async Task ExecuteSearchAsync_Preserves_PaginationFiltersAndResultStructure()
        {
            var mockLicensing = new MockLicensingService(LicenseStatus.Valid);
            var service = new SearchService(new MockSearchEngine(), new MockRepository(), mockLicensing);

            var filters = new SearchQuery
            {
                Limit = 25,
                Offset = 50,
                SenderFilter = "sender@example.com"
            };

            var result = await service.ExecuteSearchAsync("pagination test", filters);

            Assert.NotNull(result);
            Assert.Equal(100, result.TotalCount);
            Assert.True(result.ElapsedSeconds >= 0);
            Assert.Equal(1, mockLicensing.GetStatusCallCount);
        }
    }
}
