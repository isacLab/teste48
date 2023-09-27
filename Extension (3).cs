using myLIMSweb.Business.v3.Utl;
using myLIMSweb.Data;
using myLIMSweb.DataAccess.v2;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;

namespace myLIMSweb.Extensions.Task.v3
{
    public class Extension : IExtension
    {
        #region Parametros
        private Business.v3.IContext BusinessContext;
        private IContext DataContext;

        private List<string> messages = new List<string>();

        private int? sampleId;
        private int equivalencyId;
        private int workMasterId;
        private int skipLoteSampleQty;
        private int workInitialSetpId;
        private int workFinalSetpId;
        private string sampleIdentificationPrefix;
        private int? mailingListId;
        private string mailFrom;
        private int? messageTypeId;
        private int? sampleTypeId;

        private Sample sample;
        private int workId;
        private int workSampleCount;
        #endregion

        public ExecuteReturn Execute(ExecuteParameters executeParameters)
        {
            try
            {
                LoadParameters(executeParameters);

                if (sampleId != null)
                {
                    if (!IsSampleReceived())
                        return null;

                    if (!IsSkipLoteSample())
                        return null;

                    GetSample();
                    GetWorkId();

                    if (workId != 0)
                    {
                        messages.Add("Atividade encontrada Id: " + workId.ToString());
                        GetWorkSampleCount();

                        if (workSampleCount <= skipLoteSampleQty)
                        {
                            messages.Add("Amostras vinculadas: " + workSampleCount.ToString());
                            if (IsSampleAttached())
                                return null;

                            if (workSampleCount + 1 >= skipLoteSampleQty)
                            {
                                var result = AdvanceWorkFlowStep();
                                if (result.Success != true)
                                {
                                    SendWarningMessage(result);
                                    return null;
                                }

                                BusinessContext.Works.AttachSample(workId, (int)sampleId);
                                SendSampleFullyRelizedMessage();
                                ChangeSampleIdentification();
                                AddAnalysesGroupAnalyses();
                            }
                            else
                            {
                                BusinessContext.Works.AttachSample(workId, (int)sampleId);
                            }
                        }
                        else
                        {
                            CreateWork(sample);
                        }
                    }
                    else
                    {
                        CreateWork(sample);
                    }

                    this.LogNewEvent( new LogEvent
                    {
                        Level = Severity.Informational.ToString(),
                        Name = "SkipLote",
                        Source = "myLIMSweb.Task.SantaClaraAgro.SkipLote",
                        Message = "Task Executed Sucessfully.",
                        FriendlyMessage = string.Join("<br>", messages),
                        EventDateTime = DateTime.UtcNow
                    });

                }
            }
            catch (Exception ex)
            {
                this.LogNewEvent(new LogEvent
                {
                    Level = Severity.Error.ToString(),
                    Name = "SkipLote",
                    Source = "myLIMSweb.Task.SantaClaraAgro.SkipLote",
                    StackTrace = ex.StackTrace,
                    Message = GetFullException(ex),
                    FriendlyMessage = "Error When Executing Task.",
                    EventDateTime = DateTime.UtcNow
                });
            }
            return null;
        }

        private void AddAnalysesGroupAnalyses()
        {
            var analysisGroupId = DataContext.EquivalencySampleTypes
                .Where(x => x.EquivalencyId == equivalencyId
                    && x.SampleTypeId == sampleTypeId
                    && x.Equivalency.Active)
                .Select(x => x.ExternalId)
                .FirstOrDefault();

            if (!string.IsNullOrEmpty(analysisGroupId) && int.TryParse(analysisGroupId, out int numericValue))
            {
                BusinessContext.Samples.AddAnalysesByAnalysisGroup((int)sampleId, numericValue);
                DataContext.SaveChanges();
            }
        }

        private void SendSampleFullyRelizedMessage()
        {
            var subject = "Amostra #" + sampleId + " - " + sampleIdentificationPrefix;
            var message = "Amostra de ID " + sampleId + " deverá ser realizada por completo.";

            SendMessage(sampleId, subject, message);
        }

        private void ChangeSampleIdentification()
        {
            sample.Identification = sampleIdentificationPrefix + sample.Identification;
            DataContext.SaveChanges();
        }

        private void SendWarningMessage(DTOs.DefaultResult result)
        {
            messages.Add("Erro ao finalizar atividade: " + workId.ToString() + " - " + string.Join("; ", result.Messages));

            var warningSubject = "Amostra #" + sampleId + " - Problema ao finalizar Atividade #" + workId;
            var warningMessage = string.Format("Atividade Id: {0} não pôde ser finalizada.<br><br>Erro: {1}<br><br>Corrija o problema e execute a tarefa manualmente.", workId, string.Join("; ", result.Messages));

            SendMessage(sampleId, warningSubject, warningMessage);
        }

        private DTOs.DefaultResult AdvanceWorkFlowStep()
        {
            var data = BusinessContext.Works.GetDataForNextStep(workId, workFinalSetpId);
            var result = BusinessContext.Works.WorkFlowNextStepToId(workId, workFinalSetpId, data.Data);
            return result;
        }

        private void GetWorkSampleCount()
        {
            workSampleCount = DataContext.WorkSamples
                .Where(x => x.WorkId == workId
                    && x.Sample.Active
                    && !x.Sample.Reviewed)
                .Count();
        }

        private bool IsSampleAttached()
        {
            return DataContext.WorkSamples
               .Any(x => x.WorkId == workId
                   && x.SampleId == sampleId
                   && x.Sample.Active
                   && !x.Sample.Reviewed);
        }

        private void GetWorkId()
        {
            workId = DataContext.Works
                 .AsNoTracking()
                 .Where(x => x.WorkType.MasterId == workMasterId
                     && x.Identification == sample.SampleType.Identification
                     && x.FinishDateTime == null
                     && x.Active
                     && x.CurrentWorkFlow.WorkFlowStepToId != workFinalSetpId)
                 .Select(x => x.Id)
                 .FirstOrDefault();
        }

        private void GetSample()
        {
            sample = DataContext.Samples
                .Include("SampleType")
                .Where(x => x.Id == sampleId)
                .FirstOrDefault();

            messages.Add("Amostra Id: " + sampleId.ToString() + "<br>Tipo de Amostra: " + sample.SampleType.Identification);
        }

        private bool IsSkipLoteSample()
        {
            sampleTypeId = DataContext.Samples
                .AsNoTracking()
                .Where(x => x.Id == sampleId)
                .Select(x => x.SampleTypeId)
                .FirstOrDefault();

            var skipLoteSampleType = DataContext.EquivalencySampleTypes
                .Any(x => x.EquivalencyId == equivalencyId
                    && x.SampleTypeId == sampleTypeId
                    && x.Equivalency.Active);

            return skipLoteSampleType;
        }

        private bool IsSampleReceived()
        {
            return DataContext.Samples
                .AsNoTracking()
                .Any(x => x.Id == sampleId
                    && x.Received);
        }

        private void LoadParameters(ExecuteParameters executeParameters)
        {
            BusinessContext = executeParameters.BusinessContext;
            DataContext = BusinessContext.DataContext;

            dynamic parameters = executeParameters.Content;
            dynamic configs = executeParameters.Config != null ? JsonConvert.DeserializeObject<dynamic>(executeParameters.Config) : null;

            try { sampleId = parameters.SampleId; }
            catch { };


            if (configs.EquivalencyId == null)
                throw new Exception("Parâmetro \"EquivalencyId\" não configurado.");
            equivalencyId = (int)configs.EquivalencyId.Value;

            if (configs.SkipLoteSampleQty == null)
                throw new Exception("Parâmetro \"SkipLoteSampleQty\" não configurado.");
            skipLoteSampleQty = (int)configs.SkipLoteSampleQty.Value;

            if (skipLoteSampleQty <= 1)
                throw new Exception("Parâmetro \"SkipLoteSampleQty\" deve ser maior que 1.");

            if (configs.WorkFinalSetpId == null)
                throw new Exception("Parâmetro \"WorkFinalSetpId\" não configurado.");
            workFinalSetpId = (int)configs.WorkFinalSetpId.Value;

            if (configs.WorkInitialSetpId == null)
                throw new Exception("Parâmetro \"WorkInitialSetpId\" não configurado.");
            workInitialSetpId = (int)configs.WorkInitialSetpId.Value;

            if (configs.WorkMasterId == null)
                throw new Exception("Parâmetro \"WorkMasterId\" não configurado.");
            workMasterId = (int)configs.WorkMasterId.Value;

            if (configs.SampleIdentification == null)
                throw new Exception("Parâmetro \"SampleIdentification\" não configurado.");
            sampleIdentificationPrefix = (string)configs.SampleIdentification.Value;

            mailingListId = (int?)configs.MailingListId.Value;

            if (mailingListId != null && configs.MailFrom == null)
                throw new Exception("Parâmetro \"MailFrom\" não configurado.");
            mailFrom = (string)configs.MailFrom.Value;

            if (mailingListId != null && configs.MessageTypeId == null)
                throw new Exception("Parâmetro \"MessageTypeId\" não configurado.");
            messageTypeId = (int?)configs.MessageTypeId.Value;
        }

        private void SendMessage(int? sampleId, string subject, string message)
        {
            if (mailingListId != null && mailingListId > 0)
            {
                var recipientsList = DataContext.MailingListRecipients
                        .Where(x =>
                            x.MailingListId == mailingListId
                        )
                        .ToList();

                foreach (var recipient in recipientsList)
                {
                    var accountId = DataContext.AccountEmails.Where(x => x.Email == recipient.Email).Select(x => x.AccountId).FirstOrDefault();

                    var htmlBody = message;
                    var subjectString = subject;

                    List<myLIMSweb.Data.MessageRecipient> messageUsers = new List<Data.MessageRecipient>();
                    myLIMSweb.Data.MessageRecipient messageUser = new Data.MessageRecipient();
                    messageUser.AccountToId = accountId != 0 ? accountId : recipient.AccountId;
                    messageUser.Email = recipient.Email != null ? recipient.Email : null;
                    messageUsers.Add(messageUser);

                    CreateMessage(mailFrom, messageUsers, subjectString, htmlBody, messageTypeId, (int)sampleId, "");
                }
            }
            else
            {
                var newMessage = BusinessContext.Messages.New();

                newMessage.MessageHtml = message;
                newMessage.Subject = subject;

                BusinessContext.Samples.AddSampleMessage((int)sampleId, newMessage);
            }
        }

        private Data.Message CreateMessage(string from, ICollection<Data.MessageRecipient> to, string subject, string htmlbody, int? messageTypeId, int sampleId, string identifier = null)
        {
            string fileIdentification = "Mensagem automática enviada através de Tarefa Personalizada";
            string category = "html";

            if (htmlbody == null) htmlbody = String.Empty;

            byte[] bytes = System.Text.Encoding.Default.GetBytes(htmlbody);
            var file = new myLIMSweb.DTOs.FileNew();
            file.Identification = fileIdentification;
            file.Category = category;
            file.FileData = bytes;
            var messageHtmlFileId = BusinessContext.Files.Insert(file);

            Data.Message Message = new Data.Message()
            {
                Active = true,
                Draft = false,
                Identifier = identifier,
                EmailFrom = from,
                MessageHtmlFileId = messageHtmlFileId,
                MessageTextPlan = HTMLUtils.HTMLStringToPlaintText(htmlbody),
                MessageID = Guid.NewGuid().ToString(),
                MessageTypeId = (int)messageTypeId,
                Sent = DateTime.UtcNow,
                Subject = subject,
                MessageHtml = htmlbody
            };
            DataContext.Messages.Add(Message);
            DataContext.SaveChanges();

            foreach (var Recipient in to)
            {
                Recipient.MessageId = Message.Id;
                Recipient.MessageDate = DateTime.UtcNow;
                DataContext.MessageRecipients.Add(Recipient);
                DataContext.SaveChanges();
            }

            Data.SampleMessage sampleMessage = new Data.SampleMessage()
            {
                MessageId = Message.Id,
                SampleId = sampleId,
                Message = Message
            };
            DataContext.SampleMessages.Add(sampleMessage);
            DataContext.SaveChanges();

            return Message;
        }

        private void CreateWork(Sample sample)
        {
            var newWork = BusinessContext.Works.New(workMasterId, workInitialSetpId);

            newWork.Identification = sample.SampleType.Identification;

            var newWorkId = BusinessContext.Works.Insert(newWork);

            messages.Add("Nova atividade Id: " + newWorkId.ToString());

            BusinessContext.Works.AttachSample(newWorkId, sample.Id);
        }

        private void LogNewEvent(LogEvent logEvent)
        {
            DataContext.LogEvents.Add(logEvent);
            DataContext.SaveChanges();
        }

        private static string GetFullException(Exception ex)
        {
            if (ex.InnerException != null)
            {
                return ex.Message + " " + GetFullException(ex.InnerException);
            }
            else
            {
                return ex.Message;
            }
        }

        private enum Severity
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