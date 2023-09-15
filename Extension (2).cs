using System;
using System.Collections.Generic;
using System.Linq;
using myLIMSweb.DataAccess.v2;
using Newtonsoft.Json;
using System.Text;
using myLIMSweb.Task.DTOs;
using System.Globalization;
using myLIMSweb.Business.v3.Utils;
using myLIMSweb.Data;
using DevExpress.Pdf;
using System.IO;
using myLIMSweb.DTOs;

namespace myLIMSweb.Extensions.Task.v3
{
    public class Extension : IExtension
    {
        #region Parâmetros
        public Business.v3.IContext BusinessContext;
        public myLIMSweb.DataAccess.v2.IContext DataContext;

        public int workId;
        public string action = "";
        public myLIMSweb.Task.DTOs.Work workDetails;
        public string sampleFileIdentification;
      
        List<int> sampleIds = new List<int>();
        #endregion

        public ExecuteReturn Execute(ExecuteParameters parameters)
        {
            try
            {
                InitializeExecute(parameters);

                GetWorkDetails();

                ProcessWork();

            }
            catch (Exception ex)
            {
                LogErrorMessage(ex);
            }

            return null;
        }

        private void ProcessWork()
        {
            var sampleIds = GetRelatedSampleIds();

            if (sampleIds.Any())
            {
                ProcessFile(sampleIds);
            }
        }

        private void ProcessFile(List<int> sampleIds)
        {
            List<int> fileIds = GetFileIds(sampleIds);

            if (fileIds.Any())
            {
                var mergedFile = MergePDFs(fileIds);

                int fileId = InsertFile(mergedFile);

                RegisterWorkFile(fileId);

                DataContext.SaveChanges();

                LogSucessMessage();
            }
        }

        private void RegisterWorkFile(int fileId)
        {
            var workFile = new Data.WorkFile
            {
                WorkId = workId,
                FileId = fileId,
                EditionDateTime = DateTime.UtcNow,
                SyncPortal = false,
                EditionUserId = BusinessContext.AuthenticatedUser.Id,
                Active = true,
            };

            DataContext.WorkFiles.Add(workFile);
        }

        private int InsertFile(byte[] mergedFile)
        {
            return BusinessContext.Files.Insert(new FileNew
            {
                FileData = mergedFile.ToArray(),
                Identification = $"{workDetails.WorkTypeIdentification} - {workDetails.ControlNumber}.pdf",
                Category = null,
                AllowDownload = true,
                DisplayVersion = "",
            });
        }

        private byte[] MergePDFs(List<int> fileIds)
        {
            List<byte[]> pdfBytes = new List<byte[]>();
            foreach (var file in fileIds)
            {
                pdfBytes.Add(GetFileData(file));

            }
            return CombinePDFs(pdfBytes).ToArray();
        }

        private List<int> GetFileIds(List<int> sampleIds)
        {
            return DataContext.SampleFiles
                            .Where(x => sampleIds.Contains(x.SampleId) && x.File.Identification.Contains(sampleFileIdentification))
                            .OrderBy(x => x.SampleId)
                            .Select(x => x.FileId)
                            .ToList();
        }

        private List<int> GetRelatedSampleIds()
        {
            return DataContext.WorkSamples
                .Where(x => x.WorkId == workId
                    && x.Sample.Published
                    && !x.Sample.Reviewed
                    && x.Sample.Active)
                .Select(x => x.SampleId)
                .ToList();
        }

        private void GetWorkDetails()
        {
            workDetails = DataContext.Works
                .Where(x => x.Id == workId)
                .Select(x => new myLIMSweb.Task.DTOs.Work
                {
                    Id = x.Id,
                    ControlNumber = x.ControlNumber,
                    WorkTypeIdentification = x.WorkType.Identification
                })
                .FirstOrDefault();
        }

        public static MemoryStream CombinePDFs(List<byte[]> pdfBytesList)
        {
            using (PdfDocumentProcessor pdfDocumentProcessor = new PdfDocumentProcessor())
            {
                pdfDocumentProcessor.CreateEmptyDocument();
                foreach (byte[] pdfBytes in pdfBytesList)
                {
                    pdfDocumentProcessor.AppendDocument(new MemoryStream(pdfBytes));
                }

                using (MemoryStream mergedPdfStream = new MemoryStream())
                {
                    pdfDocumentProcessor.SaveDocument(mergedPdfStream);

                    mergedPdfStream.Seek(0, SeekOrigin.Begin);

                    return mergedPdfStream;
                }
            }

        }

        private void InitializeExecute(ExecuteParameters parameters)
        {
            action = "Execute";
            BusinessContext = parameters.BusinessContext;
            DataContext = BusinessContext.DataContext;
            LoadParameters(parameters);
        }

        public void LoadParameters(ExecuteParameters parameters)
        {
            action = "LoadParameters";
            dynamic content = parameters.Content;
            dynamic configs = parameters.Config != null ? JsonConvert.DeserializeObject<dynamic>(parameters.Config) : null;

            AssignParametersFromContent(content, configs);
        }

        private void AssignParametersFromContent(dynamic content, dynamic configs)
        {
            workId = content.WorkId;
            sampleFileIdentification = configs.SampleFileIdentification;
        }

        private byte[] GetFileData(int fileId)
        {
            action = $"GetFileData fileId: {fileId}";
            return BusinessContext.Files.GetFileData(fileId);
        }

        private void LogSucessMessage()
        {
            LogNewEvent(new LogEvent
            {
                Level = Severity.Informational.ToString(),
                Name = "SpreadsheetAnalysisIntegration",
                Source = "myLIMSweb.Task.ITP.ConcatenateSampleAnalysesReport",
                StackTrace = null,
                Message = $"Tarefa Executada com sucesso.",
                FriendlyMessage = $"Tarefa Executada com sucesso. WorkId: {workId.ToString()}",
                EventDateTime = DateTime.UtcNow
            });
        }

        private void LogErrorMessage(Exception ex)
        {
            LogNewEvent(new LogEvent
            {
                Level = Severity.Error.ToString(),
                Name = "ConcatenateSampleAnalysesReport",
                Source = "myLIMSweb.Task.ITP.ConcatenateSampleAnalysesReport",
                StackTrace = ex.StackTrace,
                Message = $"Ação que gerou a exceção: {action} <br>{GetFullException(ex)}",
                FriendlyMessage = $"Erro ao executar tarefa. WorkId: {workId.ToString()}",
                EventDateTime = DateTime.UtcNow
            });
        }

        private void LogNewEvent(LogEvent logEvent)
        {
            DataContext.LogEvents.Add(logEvent);
            DataContext.SaveChanges();
        }

        private string GetFullException(Exception ex)
        {
            return ex.InnerException != null ? ex.Message + " " + GetFullException(ex.InnerException) : ex.Message;
        }

        public enum Severity
        {
            Alert,
            Critical,
            Error,
            Warning,
            Notice,
            Informational,
            Debug,
            Unhandled
        }
    }
}