using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using PrintServices.DAL;
using System.Collections.Generic;
using PrintServices.Models;
using Moq;
using System.Data.Entity;
using System.Linq;

namespace PrintServices.DAL.Tests
{
    [TestClass()]
    public class RepoTests
    {
        private Mock<JobContext> mock_context { get; set; }
        private JobRepo repo { get; set; }
        private Mock<DbSet<Job>> mock_jobs { get; set; }
        private List<Job> Jobs { get; set; }

        public void ConnectToDatastore()
        {
            var query_jobs = Jobs.AsQueryable();

            mock_jobs.As<IQueryable<Job>>().Setup(m => m.Provider).Returns(query_jobs.Provider);
            mock_jobs.As<IQueryable<Job>>().Setup(m => m.Expression).Returns(query_jobs.Expression);
            mock_jobs.As<IQueryable<Job>>().Setup(m => m.ElementType).Returns(query_jobs.ElementType);
            mock_jobs.As<IQueryable<Job>>().Setup(m => m.GetEnumerator()).Returns(() => query_jobs.GetEnumerator());

            mock_context.Setup(c => c.Jobs).Returns(mock_jobs.Object);
            mock_jobs.Setup(u => u.Add(It.IsAny<Job>())).Callback((Job t) => Jobs.Add(t));
            mock_jobs.Setup(u => u.Remove(It.IsAny<Job>())).Callback((Job t) => Jobs.Remove(t));
        }

        [TestInitialize]
        public void Initialize()
        {
            mock_jobs = new Mock<DbSet<Job>>();
            Jobs = new List<Job>();

            mock_context = new Mock<JobContext>();
            repo = new JobRepo(mock_context.Object);

            ConnectToDatastore();
        }
        public void ImportMockData()
        {
            repo.ImportMasterSpreadsheet();
        }
        
        [TestCleanup]
        public void TearDown()
        {
            repo = null;
        }

        [TestMethod]
        public void CanInstantiateRepo()
        {
            Assert.IsNotNull(repo);
        }

        [TestMethod]
        public void CanTearDownRepo()
        {
            TearDown();
            Assert.IsNull(repo);
        }
        [TestMethod]
        public void CanImportMockData()
        {
            ImportMockData();

            Assert.IsTrue(Jobs.Count > 100);
        }


    }
}
