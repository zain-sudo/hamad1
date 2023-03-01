// <copyright file="EventGraphHelper.cs" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

namespace Microsoft.Teams.Apps.EmployeeTraining.Helpers
{
    extern alias BetaLib;

    using System;
    using System.Collections.Generic;
    using System.Globalization;
    using System.Linq;
    using System.Net.Http.Headers;
    using System.Threading.Tasks;
    using System.Web;
    using Microsoft.ApplicationInsights;
    using Microsoft.AspNetCore.Http;
    using Microsoft.Exchange.WebServices.Data;
    using Microsoft.Extensions.Localization;
    using Microsoft.Extensions.Options;
    using Microsoft.Graph;
    using Microsoft.Teams.Apps.EmployeeTraining.Models;
    using Microsoft.Teams.Apps.EmployeeTraining.Models.Configuration;
#pragma warning disable SA1135 // Referring BETA package of MS Graph SDK.
    using Beta = BetaLib.Microsoft.Graph;
#pragma warning restore SA1135 // Referring BETA package of MS Graph SDK.
    using EventType = Microsoft.Teams.Apps.EmployeeTraining.Models.EventType;

    /// <summary>
    /// Implements the methods that are defined in <see cref="IEventGraphHelper"/>.
    /// </summary>
    public class EventGraphHelper : IEventGraphHelper
    {
        /// <summary>
        /// Instance service email;
        /// </summary>
        private readonly string serviceEmail;

        /// <summary>
        /// Instance service password;
        /// </summary>
        private readonly string servicePass;

        /// <summary>
        /// Instance EWS URL;
        /// </summary>
        private readonly string ewsUrl;

        /// <summary>
        /// Represents a set of key/value application configuration properties for Azure.
        /// </summary>
        private readonly IOptions<AzureVaultSettings> azureVaultOptions;

        /// <summary>
        /// Instance of graph service client for delegated requests.
        /// </summary>
        private readonly GraphServiceClient delegatedGraphClient;

        /// <summary>
        /// Instance of graph service client for application level requests.
        /// </summary>
        private readonly GraphServiceClient applicationGraphClient;

        /// <summary>
        /// Instance of BETA graph service client for application level requests.
        /// </summary>
        private readonly Beta.GraphServiceClient applicationBetaGraphClient;

        /// <summary>
        /// The current culture's string localizer
        /// </summary>
        private readonly IStringLocalizer<Strings> localizer;

        /// <summary>
        /// Graph helper for operations related user.
        /// </summary>
        private readonly IUserGraphHelper userGraphHelper;

        /// <summary>
        /// Instance onPremises user;
        /// </summary>
        private bool isOnPremUser;

        /// <summary>
        /// Instance userName;
        /// </summary>
        private string userName;

        /// <summary>
        /// Initializes a new instance of the <see cref="EventGraphHelper"/> class.
        /// </summary>
        /// <param name="tokenAcquisitionHelper">Helper to get user access token for specified Graph scopes.</param>
        /// <param name="httpContextAccessor">HTTP context accessor for getting user claims.</param>
        /// <param name="localizer">The current culture's string localizer.</param>
        /// <param name="userGraphHelper">Graph helper for operations related user.</param>
        /// <param name="azureVaultOptions">A set of key/value application configuration properties for Key Vault.</param>
        public EventGraphHelper(
            ITokenAcquisitionHelper tokenAcquisitionHelper,
            IHttpContextAccessor httpContextAccessor,
            IStringLocalizer<Strings> localizer,
            IUserGraphHelper userGraphHelper,
            IOptions<AzureVaultSettings> azureVaultOptions)
        {
            this.localizer = localizer;
            this.userGraphHelper = userGraphHelper;
            httpContextAccessor = httpContextAccessor ?? throw new ArgumentNullException(nameof(httpContextAccessor));
            this.azureVaultOptions = azureVaultOptions ?? throw new ArgumentNullException(nameof(azureVaultOptions));
            this.serviceEmail = this.azureVaultOptions.Value.ServiceEmail;
            this.servicePass = this.azureVaultOptions.Value.ServicePassword;
            this.ewsUrl = this.azureVaultOptions.Value.EwsUrl;

            var oidClaimType = "http://schemas.microsoft.com/identity/claims/objectidentifier";
            var userObjectId = httpContextAccessor.HttpContext.User.Claims?
                .FirstOrDefault(claim => oidClaimType.Equals(claim.Type, StringComparison.OrdinalIgnoreCase))?.Value;

            if (!string.IsNullOrEmpty(userObjectId))
            {
                var jwtToken = AuthenticationHeaderValue.Parse(httpContextAccessor.HttpContext.Request.Headers["Authorization"].ToString()).Parameter;

                this.delegatedGraphClient = GraphServiceClientFactory.GetAuthenticatedGraphClient(async () =>
                {
                    return await tokenAcquisitionHelper.GetUserAccessTokenAsync(userObjectId, jwtToken);
                });

                this.applicationBetaGraphClient = GraphServiceClientFactory.GetAuthenticatedBetaGraphClient(async () =>
                {
                    return await tokenAcquisitionHelper.GetApplicationAccessTokenAsync();
                });

                this.applicationGraphClient = GraphServiceClientFactory.GetAuthenticatedGraphClient(async () =>
                {
                    return await tokenAcquisitionHelper.GetApplicationAccessTokenAsync();
                });

                this.isOnPremUser = this.delegatedGraphClient.Me.Request().Select("onPremisesSyncEnabled").GetAsync().Result.OnPremisesSyncEnabled.HasValue;
                this.userName = this.delegatedGraphClient.Me.Request().Select("userPrincipalName").GetAsync().Result.UserPrincipalName;
            }
        }

        /// <summary>
        /// Instance create appointemnt or delete appointment;
        /// </summary>
        private enum CreateUpdate
        {
            CreateAppointment,
            UpdateAppointment,
        }

        /// <summary>
        /// Cancel calendar event.
        /// </summary>
        /// <param name="eventGraphId">Event Id received from Graph.</param>
        /// <param name="createdByUserId">User Id who created event.</param>
        /// <param name="comment">Cancellation comment.</param>
        /// <param name="telemetryClient">telemetry</param>
        /// <returns>True if event cancellation is successful.</returns>
        public async Task<bool> CancelEventAsync(string eventGraphId, string createdByUserId, string comment, TelemetryClient telemetryClient)
        {
            telemetryClient?.TrackTrace($"{this.userName} EventGraphHelper Event:CancelEventAsync CALLED");

            try
            {
                telemetryClient?.TrackEvent($"{this.userName} EventGraphHelper Event:CancelEventAsync:Canceling event:{eventGraphId}");

                Item item;
                ExchangeService service;

                ItemId eventId = eventGraphId;
                var user = await this.delegatedGraphClient.Users[createdByUserId].Request().GetAsync();
                string userPrincipal = user.UserPrincipalName;
                bool isCreatorOnPrem = this.delegatedGraphClient.Users[createdByUserId].Request().Select("onPremisesSyncEnabled").GetAsync().Result.OnPremisesSyncEnabled.HasValue;
                if (isCreatorOnPrem)
                {
                    telemetryClient?.TrackEvent($"{this.userName} EventGraphHelper Event:CancelEventAsync:On-Prem event Cancelling:{eventGraphId}");
                    service = this.EwsService(userPrincipal, telemetryClient);
                    item = Item.Bind(service, eventId);
                    item.Delete(DeleteMode.MoveToDeletedItems);
                    telemetryClient?.TrackTrace($"{this.userName} EventGraphHelper Event:CancelEventAsync:On-Prem event Cancelled:{eventGraphId}");
                    telemetryClient?.TrackTrace($"{this.userName} EventGraphHelper Event:CancelEventAsync SUCCESS");
                    return true;
                }
                else
                {
                    telemetryClient?.TrackEvent($"{this.userName} EventGraphHelper Event:CancelEventAsync:On-line event Cancelling:{eventGraphId}");
                    await this.applicationBetaGraphClient.Users[createdByUserId].Events[eventGraphId].Cancel(comment).Request().PostAsync();
                    telemetryClient?.TrackTrace($"{this.userName} EventGraphHelper Event:CancelEventAsync:On-line event Cancelled:{eventGraphId}");
                    telemetryClient?.TrackTrace($"{this.userName} EventGraphHelper Event:CancelEventAsync SUCCESS");
                    return true;
                }
            }
            catch (Exception ex)
            {
                var ewsExcption = new Exception($"{this.userName} EventGraphHelper Event:CreateEventAsync: FAILED with Exception:{ex.Message}");
                telemetryClient.TrackException(ewsExcption);
                throw new Exception(ewsExcption.Message);
            }
        }

        /// <summary>
        /// Create teams event.
        /// </summary>
        /// <param name="eventEntity">Event details from user for which event needs to be created.</param>
        /// /// <param name="telemetryClient">telemetry</param>
        /// <returns>Created event details.</returns>
        public async Task<Event> CreateEventAsync(EventEntity eventEntity, TelemetryClient telemetryClient)
        {
            telemetryClient?.TrackTrace($"{this.userName} EventGraphHelper Event:CreateEventAsync CALLED");
            try
            {
                telemetryClient?.TrackEvent($"{this.userName} EventGraphHelper Event:CreateEventAsync:Creating event:{eventEntity?.Name}");
                eventEntity = eventEntity ?? throw new ArgumentNullException(nameof(eventEntity), "Event details cannot be null");

                var teamsEvent = new Event { };
                ExchangeService service;

                string userPrincipal = this.userName;
                teamsEvent.Subject = eventEntity.Name;
                teamsEvent.Body = new ItemBody
                {
                    ContentType = Microsoft.Graph.BodyType.Html,
                    Content = await this.GetEventBodyContent(eventEntity, telemetryClient, CreateUpdate.CreateAppointment),
                };
                teamsEvent.Attendees = eventEntity.IsAutoRegister && eventEntity.Audience == (int)EventAudience.Private ?
                            await this.GetEventAttendeesTemplateAsync(eventEntity, telemetryClient) :
                            new List<Microsoft.Graph.Attendee>();
                teamsEvent.OnlineMeetingUrl = eventEntity.Type == (int)EventType.LiveEvent ? eventEntity.MeetingLink : null;
                teamsEvent.IsReminderOn = true;
                teamsEvent.Location = eventEntity.Type == (int)EventType.InPerson ? new Location
                {
                    DisplayName = eventEntity.Venue,
                }
                : null;
                teamsEvent.AllowNewTimeProposals = false;
                teamsEvent.IsOnlineMeeting = eventEntity.Type == (int)EventType.Teams;
                teamsEvent.OnlineMeetingProvider = eventEntity.Type == (int)EventType.Teams ? OnlineMeetingProviderType.TeamsForBusiness : OnlineMeetingProviderType.Unknown;
                teamsEvent.Start = new DateTimeTimeZone
                {
                    DateTime = eventEntity.StartDate?.ToString("s", CultureInfo.InvariantCulture),
                    TimeZone = TimeZoneInfo.Utc.Id,
                };
                teamsEvent.End = new DateTimeTimeZone
                {
                    // DateTime = eventEntity.StartDate.Value.Date.Add(
                    // new TimeSpan(eventEntity.EndTime.Hour, eventEntity.EndTime.Minute, eventEntity.EndTime.Second)).ToString("s", CultureInfo.InvariantCulture),
                    // DateTime = eventEntity.EndDate?.ToString("s", CultureInfo.InvariantCulture),
                    DateTime = this.SingleEventEndDate((DateTime)eventEntity.StartDate, (DateTime)eventEntity.EndDate, telemetryClient).ToString("s", CultureInfo.InvariantCulture),
                    TimeZone = TimeZoneInfo.Utc.Id,
                };

                teamsEvent.Recurrence = null;
                if (eventEntity.NumberOfOccurrences > 1)
                {
                    // Create recurring event.
                    teamsEvent = this.GetRecurringEventTemplate(teamsEvent, eventEntity, telemetryClient);
                }

                if (this.isOnPremUser)
                {
                    telemetryClient?.TrackEvent($"{this.userName} EventGraphHelper Event:CreateEventAsync:On-prem event:{eventEntity?.Name}  Meeting Creating");

                    service = this.EwsService(userPrincipal, telemetryClient);
                    telemetryClient?.TrackEvent($"{this.userName} EventGraphHelper Event:CreateEventAsync:On-prem Creating:{eventEntity?.Name}");
                    this.CreateEWSEvent(telemetryClient, service, teamsEvent);
                    telemetryClient?.TrackTrace($"{this.userName} EventGraphHelper Event:CreateEventAsync:On-prem Created:{eventEntity?.Name}");
                    telemetryClient?.TrackTrace($"{this.userName} EventGraphHelper Event:CreateEventAsync SUCCESS");
                    return teamsEvent;
                }
                else
                {
                    telemetryClient?.TrackEvent($"{this.userName} EventGraphHelper Event:CreateEventAsync:On-line Creating:{eventEntity?.Name}");
                    var cloudEvent = await this.delegatedGraphClient.Me.Events.Request().Header("Prefer", $"outlook.timezone=\"{TimeZoneInfo.Utc.Id}\"").AddAsync(teamsEvent);
                    telemetryClient?.TrackTrace($"{this.userName} EventGraphHelper Event:CreateEventAsync:On-line Created:{eventEntity?.Name}");
                    telemetryClient?.TrackTrace($"{this.userName} EventGraphHelper Event:CreateEventAsync SUCCESS");
                    return cloudEvent;
                }
            }
            catch (Exception ex)
            {
                var ewsExcption = new Exception($"{this.userName} EventGraphHelper Event:CreateEventAsync: FAILED with Message:{ex.Message} and Exception:{ex.StackTrace}");
                telemetryClient.TrackException(ewsExcption);
                throw new Exception(ewsExcption.Message);
            }
        }

        /// <summary>
        /// Update teams event.
        /// </summary>
        /// <param name="eventEntity">Event details from user for which event needs to be updated.</param>
        /// <param name="telemetryClient">telemetry</param>
        /// <returns>Updated event details.</returns>
        public async Task<Event> UpdateEventAsync(EventEntity eventEntity, TelemetryClient telemetryClient)
        {
            telemetryClient?.TrackTrace($"{this.userName} EventGraphHelper Event:UpdateEventAsync CALLED");
            try
            {
                telemetryClient?.TrackEvent($"{this.userName} EventGraphHelper Event:UpdateEventAsync:updating event:{eventEntity?.GraphEventId}");
                eventEntity = eventEntity ?? throw new ArgumentNullException(nameof(eventEntity), "Event details cannot be null");

                ItemId eventId = eventEntity.GraphEventId;
                var teamsEvent = new Event { };
                ExchangeService service;

                bool isCreatedByOnPremUser = this.delegatedGraphClient.Users[eventEntity.CreatedBy].Request().Select("onPremisesSyncEnabled").GetAsync().Result.OnPremisesSyncEnabled.HasValue;

                var user = await this.delegatedGraphClient.Users[eventEntity.CreatedBy].Request().GetAsync();
                string userPrincipal = user.UserPrincipalName;
                teamsEvent.Subject = eventEntity.Name;
                teamsEvent.Body = new ItemBody
                {
                    ContentType = Microsoft.Graph.BodyType.Html,
                    Content = await this.GetEventBodyContent(eventEntity, telemetryClient, CreateUpdate.UpdateAppointment),
                };
                telemetryClient.TrackEvent($"{this.userName} EventGraphHelper Event:UpdateEventAsync:updating event:{eventEntity.GraphEventId} with body:{teamsEvent.Body}");
                teamsEvent.Attendees = await this.GetEventAttendeesTemplateAsync(eventEntity, telemetryClient);

                if (eventEntity.Type == (int)EventType.LiveEvent)
                {
                    // Teams
                    teamsEvent.IsOnlineMeeting = false;
                    teamsEvent.OnlineMeetingProvider = OnlineMeetingProviderType.Unknown;

                    // Live
                    teamsEvent.OnlineMeetingUrl = eventEntity.MeetingLink;

                    // In-Person
                    teamsEvent.Location = null;
                }

                if (eventEntity.Type == (int)EventType.InPerson)
                {
                    // Teams
                    teamsEvent.IsOnlineMeeting = false;
                    teamsEvent.OnlineMeetingProvider = OnlineMeetingProviderType.Unknown;

                    // Live
                    teamsEvent.OnlineMeetingUrl = null;

                    // In-Person
                    teamsEvent.Location = new Location
                    {
                        DisplayName = eventEntity.Venue,
                    };
                }

                if (eventEntity.Type == (int)EventType.Teams)
                {
                    // Teams
                    teamsEvent.IsOnlineMeeting = true;
                    teamsEvent.OnlineMeetingProvider = OnlineMeetingProviderType.TeamsForBusiness;

                    // Live
                    teamsEvent.OnlineMeetingUrl = null;

                    // In-Person
                    teamsEvent.Location = null;
                }

                teamsEvent.IsReminderOn = true;
                teamsEvent.AllowNewTimeProposals = false;
                teamsEvent.Start = new DateTimeTimeZone
                {
                    DateTime = eventEntity.StartDate?.ToString("s", CultureInfo.InvariantCulture),
                    TimeZone = TimeZoneInfo.Utc.Id,
                };
                teamsEvent.End = new DateTimeTimeZone
                {
                    // DateTime = eventEntity.StartDate.Value.Date.Add(
                    // new TimeSpan(eventEntity.EndTime.Hour, eventEntity.EndTime.Minute, eventEntity.EndTime.Second)).ToString("s", CultureInfo.InvariantCulture),
                    // DateTime = eventEntity.EndDate?.ToString("s", CultureInfo.InvariantCulture),
                    DateTime = this.SingleEventEndDate((DateTime)eventEntity.StartDate, (DateTime)eventEntity.EndDate, telemetryClient).ToString("s", CultureInfo.InvariantCulture),
                    TimeZone = TimeZoneInfo.Utc.Id,
                };

                if (eventEntity.NumberOfOccurrences >= 1)
                {
                    teamsEvent = this.GetRecurringEventTemplate(teamsEvent, eventEntity, telemetryClient);
                }
                else
                {
                    teamsEvent.Recurrence = null;
                }

                if (isCreatedByOnPremUser)
                {
                    telemetryClient.TrackEvent($"{this.userName} EventGraphHelper Event:UpdateEventAsync:OnPrem event updating:{eventEntity.GraphEventId}");
                    service = this.EwsService(userPrincipal, telemetryClient);
                    this.UpdateEWSEvent(telemetryClient, service, teamsEvent, eventId);
                    telemetryClient.TrackTrace($"{this.userName} EventGraphHelper Event:UpdateEventAsync:OnPrem event Updated:{eventEntity.GraphEventId}");
                    telemetryClient?.TrackTrace($"{this.userName} EventGraphHelper Event:UpdateEventAsync SUCCESS");
                    return teamsEvent;
                }
                else
                {
                    telemetryClient.TrackEvent($"{this.userName} EventGraphHelper Event:UpdateEventAsync:On-line event updating:{eventEntity.GraphEventId}");
                    var cloudEvent = await this.applicationGraphClient.Users[eventEntity.CreatedBy].Events[eventEntity.GraphEventId].Request().Header("Prefer", $"outlook.timezone=\"{TimeZoneInfo.Utc.Id}\"").UpdateAsync(teamsEvent);
                    telemetryClient.TrackTrace($"{this.userName} EventGraphHelper Event:UpdateEventAsync:On-line event Updated:{eventEntity.GraphEventId}");
                    telemetryClient?.TrackTrace($"{this.userName} EventGraphHelper Event:UpdateEventAsync SUCCESS");
                    return cloudEvent;
                }
            }
            catch (Exception ex)
            {
                var ewsExcption = new Exception($"{this.userName} EventGraphHelper Event:UpdateEventAsync: FAILED with Message:{ex.Message} and Exception:{ex.StackTrace}");
                telemetryClient.TrackException(ewsExcption);
                throw new Exception(ewsExcption.Message);
            }
        }

        /// <summary>
        /// Changes the single events end date time so that the single event's date is same as it's starting date.
        /// </summary>
        /// <param name="startDate"> Single event's starting date time </param>
        /// <param name="endDate"> Single event's ending date time </param>
        /// <param name="telemetryClient">telemetry</param>
        /// <returns> A new date time </returns>
        private DateTime SingleEventEndDate(DateTime startDate, DateTime endDate, TelemetryClient telemetryClient)
        {
            var year = startDate.Year;
            var month = startDate.Month;
            var day = startDate.Day;

            var hour = endDate.Hour;
            var minute = endDate.Minute;

            DateTime dateTime = new DateTime(year, month, day, hour, minute, 0);

            telemetryClient.TrackEvent($"End-Date: {dateTime}");

            return dateTime;
        }

        /// <summary>
        /// Modify event details for recurring event creation.
        /// </summary>
        /// <param name="teamsEvent">Event details which will be sent to Graph API.</param>
        /// <param name="eventEntity">Event details from user for which event needs to be created.</param>
        /// <param name="telemetryClient">Telementry.</param>
        /// <returns>Event details to be sent to Graph API.</returns>
        private Event GetRecurringEventTemplate(Event teamsEvent, EventEntity eventEntity, TelemetryClient telemetryClient)
        {
            telemetryClient?.TrackTrace($"{this.userName} EventGraphHelper Event:GetRecurringEventTemplate CALLED");
            try
            {
                telemetryClient.TrackEvent($"{this.userName} EventGraphHelper Event:GetRecurringEventTemplate:Getting recurring event for:{eventEntity.Name}");

                // Create recurring event.
                teamsEvent.Recurrence = new PatternedRecurrence
                {
                    Pattern = new RecurrencePattern
                    {
                        Type = RecurrencePatternType.Daily,
                        Interval = 1,
                    },
                    Range = new RecurrenceRange
                    {
                        Type = RecurrenceRangeType.EndDate,
                        EndDate = new Date((int)eventEntity.EndDate?.Year, (int)eventEntity.EndDate?.Month, (int)eventEntity.EndDate?.Day),
                        StartDate = new Date((int)eventEntity.StartDate?.Year, (int)eventEntity.StartDate?.Month, (int)eventEntity.StartDate?.Day),
                        NumberOfOccurrences = eventEntity.NumberOfOccurrences,
                    },
                };

                telemetryClient?.TrackTrace($"{this.userName} EventGraphHelper Event:GetRecurringEventTemplate SUCCESS");
                return teamsEvent;
            }
            catch (Exception ex)
            {
                var ewsExcption = new Exception($"{this.userName} EventGraphHelper Event:GetRecurringEventTemplate: FAILED with Message:{ex.Message} and Exception:{ex.StackTrace}");
                telemetryClient.TrackException(ewsExcption);
                throw new Exception(ewsExcption.Message);
            }
        }

        /// <summary>
        /// Get list of event attendees for creating teams event.
        /// </summary>
        /// <param name="eventEntity">Event details containing registered attendees.</param>
        /// <param name="telemetryClient">Telementry.</param>
        /// <returns>List of attendees.</returns>
        private async Task<List<Microsoft.Graph.Attendee>> GetEventAttendeesTemplateAsync(EventEntity eventEntity, TelemetryClient telemetryClient)
        {
            telemetryClient?.TrackTrace($"{this.userName} EventGraphHelper Event:GetEventAttendeesTemplateAsync CALLED");
            try
            {
                var attendees = new List<Microsoft.Graph.Attendee>();

                if (string.IsNullOrEmpty(eventEntity.RegisteredAttendees) && string.IsNullOrEmpty(eventEntity.AutoRegisteredAttendees))
                {
                    telemetryClient?.TrackTrace($"{this.userName} EventGraphHelper Event:GetEventAttendeesTemplateAsync: RegisteredAttendees and AutoRegisteredAttendees are NULL");
                    return attendees;
                }

                if (!string.IsNullOrEmpty(eventEntity.RegisteredAttendees))
                {
                    telemetryClient?.TrackEvent($"{this.userName} EventGraphHelper Event:GetEventAttendeesTemplateAsync: Finding registered attendees for:{eventEntity.Name}");
                    var registeredAttendeesList = eventEntity.RegisteredAttendees.Trim().Split(";");

                    if (registeredAttendeesList.Any())
                    {
                        var userProfiles = await this.userGraphHelper.GetUsersAsync(registeredAttendeesList);

                        foreach (var userProfile in userProfiles)
                        {
                            attendees.Add(new Microsoft.Graph.Attendee
                            {
                                EmailAddress = new Microsoft.Graph.EmailAddress
                                {
                                    Address = userProfile.UserPrincipalName,
                                    Name = userProfile.DisplayName,
                                },
                                Type = AttendeeType.Required,
                            });
                        }
                    }

                    telemetryClient?.TrackTrace($"{this.userName} EventGraphHelper Event:GetEventAttendeesTemplateAsync: Found registered attendees for:{eventEntity.Name}");
                }

                if (!string.IsNullOrEmpty(eventEntity.AutoRegisteredAttendees))
                {
                    telemetryClient?.TrackEvent($"{this.userName} EventGraphHelper Event:GetEventAttendeesTemplateAsync: Finding auto registered attendees for:{eventEntity.Name}");
                    var autoRegisteredAttendeesList = eventEntity.AutoRegisteredAttendees.Trim().Split(";");

                    if (autoRegisteredAttendeesList.Any())
                    {
                        var userProfiles = await this.userGraphHelper.GetUsersAsync(autoRegisteredAttendeesList);

                        foreach (var userProfile in userProfiles)
                        {
                            attendees.Add(new Microsoft.Graph.Attendee
                            {
                                EmailAddress = new Microsoft.Graph.EmailAddress
                                {
                                    Address = userProfile.UserPrincipalName,
                                    Name = userProfile.DisplayName,
                                },
                                Type = AttendeeType.Required,
                            });
                        }
                    }

                    telemetryClient?.TrackTrace($"{this.userName} EventGraphHelper Event:GetEventAttendeesTemplateAsync: Found auto registered attendees for:{eventEntity.Name}");
                }

                telemetryClient?.TrackTrace($"{this.userName} EventGraphHelper Event:GetEventAttendeesTemplateAsync SUCCESS");
                return attendees;
            }
            catch (Exception ex)
            {
                var ewsExcption = new Exception($"{this.userName} EventGraphHelper Event:GetEventAttendeesTemplateAsync: FAILED with Message:{ex.Message} and Exception:{ex.StackTrace}");
                telemetryClient.TrackException(ewsExcption);
                throw new Exception(ewsExcption.Message);
            }
        }

        /// <summary>
        /// Get the event body content based on event type
        /// </summary>
        /// <param name="eventEntity">The event details</param>
        /// <param name="telemetryClient">Telementry.</param>
        /// <param name="createUpdate"> Enum for whether we're creating a new event or updating the event.</param>
        /// <returns>Returns </returns>
        private async Task<string> GetEventBodyContent(EventEntity eventEntity, TelemetryClient telemetryClient, CreateUpdate createUpdate)
        {
            telemetryClient?.TrackTrace($"{this.userName} EventGraphHelper Event:GetEventBodyContent CALLED");
            try
            {
                telemetryClient?.TrackEvent($"{this.userName} EventGraphHelper Event:GetEventBodyContent: Getting description for:{eventEntity.Name}");
                switch ((EventType)eventEntity.Type)
                {
                    case EventType.InPerson:
                        return HttpUtility.HtmlEncode(eventEntity.Description);

                    case EventType.LiveEvent:
                        return $"{HttpUtility.HtmlEncode(eventEntity.Description)}<br/><br/>{this.localizer.GetString("CalendarEventLiveEventURLText", $"<a href='{eventEntity.MeetingLink}'>{eventEntity.MeetingLink}</a>")}";

                    default:
                        string teamsMeeting = await this.CreateMeeting(eventEntity, telemetryClient, createUpdate);
                        return $"{HttpUtility.HtmlEncode(eventEntity.Description)}<br/><br/>{teamsMeeting}";
                }
            }
            catch (Exception ex)
            {
                var ewsExcption = new Exception($"{this.userName} EventGraphHelper Event:GetEventBodyContent: FAILED with Message:{ex.Message} and Exception:{ex.StackTrace}");
                telemetryClient.TrackException(ewsExcption);
                throw new Exception(ewsExcption.Message);
            }
        }

        /// <summary>
        /// This function contains information about a meeting, including the URL used to join a meeting, the attendees list, and the description
        /// </summary>
        /// <param name="eventEntity">The event details</param>
        /// <param name="telemetryClient">Telementry.</param>
        /// <param name="createUpdate"> Enum for whether we're creating a new event or updating the event.</param>
        /// <returns> Online meeting link </returns>
        private async Task<string> CreateMeeting(EventEntity eventEntity, TelemetryClient telemetryClient, CreateUpdate createUpdate)
        {
            telemetryClient?.TrackTrace($"{this.userName} EventGraphHelper Event:CreateMeeting CALLED");
            try
            {
                var onlineMeeting = new OnlineMeeting
                {
                    StartDateTime = DateTimeOffset.Parse(eventEntity.StartDate?.ToString("s", CultureInfo.InvariantCulture), CultureInfo.InvariantCulture),
                    EndDateTime = DateTimeOffset.Parse(eventEntity.EndDate?.ToString("s", CultureInfo.InvariantCulture), CultureInfo.InvariantCulture),
                    Subject = eventEntity.Name,
                };

                OnlineMeeting meeting;
                if (createUpdate.Equals(CreateUpdate.CreateAppointment))
                {
                    meeting = await this.delegatedGraphClient.Me.OnlineMeetings.Request().AddAsync(onlineMeeting);
                }
                else
                {
                    try
                    {
                        meeting = await this.delegatedGraphClient.Users[eventEntity.CreatedBy].OnlineMeetings.Request().AddAsync(onlineMeeting);
                    }
                    catch
                    {
                        meeting = await this.delegatedGraphClient.Me.OnlineMeetings.Request().AddAsync(onlineMeeting);
                    }
                }

                var myDecodedString = HttpUtility.UrlDecode(meeting.JoinInformation.Content);
                myDecodedString = myDecodedString.Remove(0, 15);

                telemetryClient?.TrackEvent($"{this.userName} EventGraphHelper Event:CreateMeeting: SUCCESS");
                return myDecodedString;
            }
            catch (Exception ex)
            {
                var ewsExcption = new Exception($"{this.userName} EventGraphHelper Event:CreateMeeting: FAILED with Message:{ex.Message} and Exception:{ex.StackTrace}");
                telemetryClient.TrackException(ewsExcption);
                throw new Exception(ewsExcption.Message);
            }
        }

        /// <summary>
        /// Create teams service.
        /// </summary>
        /// <param name="userPrincipal">Email ID of the user that is currently logged in.</param>
        /// <param name="telemetryClient">Telementry.</param>
        /// <returns>Created service.</returns>
        private ExchangeService EwsService(string userPrincipal, TelemetryClient telemetryClient)
        {
            telemetryClient?.TrackTrace($"{this.userName} EventGraphHelper OnPrem Event:GetEventBodyContent CALLED");
            try
            {
                telemetryClient?.TrackEvent($"{this.userName} EventGraphHelper OnPrem Event:GetEventBodyContent: Creating EWS service");
                var ewsClient = new ExchangeService(ExchangeVersion.Exchange2013_SP1);
                ewsClient.Credentials = new WebCredentials(this.serviceEmail, this.servicePass);
                ewsClient.ImpersonatedUserId = new ImpersonatedUserId(ConnectingIdType.SmtpAddress, userPrincipal);

                ewsClient.Url = new Uri(this.ewsUrl);

                telemetryClient?.TrackTrace($"{this.userName} EventGraphHelper OnPrem Event:EwsService SUCCESS");
                return ewsClient;
            }
            catch (Exception ex)
            {
                var ewsExcption = new Exception($"{this.userName} EventGraphHelper OnPrem Event:EwsService: FAILED with Message:{ex.Message} and Exception:{ex.StackTrace}");
                telemetryClient.TrackException(ewsExcption);
                throw new Exception(ewsExcption.Message);
            }
        }

        /// <summary>
        /// Create teams event.
        /// </summary>
        /// <param name="telemetryClient">Telementry.</param>
        /// <param name="service">Exchange service that will be used to create event.</param>
        /// <param name="teamsEvent">Details that need to be filled in the event.</param>
        /// <returns>Id of the event created</returns>
        private ItemId CreateEWSEvent(TelemetryClient telemetryClient, ExchangeService service, Event teamsEvent)
        {
            telemetryClient.TrackTrace($"{this.userName} EventGraphHelper OnPrem Event:CreateEWSEvent CALLED");
            try
            {
                telemetryClient?.TrackEvent($"{this.userName} EventGraphHelper Event:GetEventBodyContent: creating new appointment for:{teamsEvent.Subject}");

                Appointment appointment;
                CreateUpdate createEvent = CreateUpdate.CreateAppointment;

                ItemId blankID = new ItemId("000");

                appointment = this.UpsertEWSAppointment(teamsEvent, createEvent, service, blankID, telemetryClient);
                Item item = Item.Bind(service, appointment.Id, new PropertySet(ItemSchema.Subject));
                teamsEvent.Id = item.Id.ToString();
                ItemId eventId = appointment.Id;
                telemetryClient.TrackTrace($"{this.userName} EventGraphHelper OnPrem Event:CreateEWSEvent:Event creation SUCCESS");
                return eventId;
            }
            catch (Exception ex)
            {
                var ewsExcption = new Exception($"{this.userName} EventGraphHelper OnPrem Event:CreateEWSEvent: FAILED with Message:{ex.Message} and Exception:{ex.StackTrace}");
                telemetryClient.TrackException(ewsExcption);
                throw new Exception(ewsExcption.Message);
            }
        }

        /// <summary>
        /// Updates the event.
        /// </summary>
        /// <param name="telemetryClient">Telementry.</param>
        /// <param name="service">Exchange service that will be used to update event.</param>
        /// <param name="teamsEvent">Details that will updated in the event</param>
        /// <param name="eventId">Id of the event that need to me modified.</param>
        private void UpdateEWSEvent(TelemetryClient telemetryClient, ExchangeService service, Event teamsEvent, ItemId eventId)
        {
            telemetryClient.TrackTrace($"{this.userName} EventGraphHelper OnPrem Event:UpdateEWSEvent CALLED");

            try
            {
                telemetryClient?.TrackEvent($"{this.userName} EventGraphHelper Event:UpdateEWSEvent: updating appointment for:{teamsEvent.Subject}");
                CreateUpdate updateEvent = CreateUpdate.UpdateAppointment;
                this.UpsertEWSAppointment(teamsEvent, updateEvent, service, eventId, telemetryClient);
                telemetryClient.TrackTrace($"{this.userName} EventGraphHelper OnPrem Event:UpdateEWSEvent: Update event SUCCESS");
            }
            catch (Exception ex)
            {
                var ewsExcption = new Exception($"{this.userName} EventGraphHelper OnPrem Event:UpdateEWSEvent: FAILED with Message:{ex.Message} and Exception:{ex.StackTrace}");
                telemetryClient.TrackException(ewsExcption);
                throw new Exception(ewsExcption.Message);
            }
        }

        /// <summary>
        /// Creates or updates an appointment.
        /// </summary>
        /// <param name="teamsEvent">Detailsof the event</param>
        /// <param name="createUpdate">Enum to check if the appointment should be created or updated</param>
        /// <param name="service">Exchange service that will be used to delete event.</param>
        /// <param name="eventId">For updating appointment an ID is required</param>
        /// <param name="telemetryClient">Telementry.</param>
        private Appointment UpsertEWSAppointment(Event teamsEvent, CreateUpdate createUpdate, ExchangeService service, ItemId eventId, TelemetryClient telemetryClient)
        {
            telemetryClient.TrackTrace($"{this.userName} EventGraphHelper OnPrem Event:UpsertEWSAppointment CALLED");
            try
            {
                Appointment appointment;
                if (createUpdate.Equals(CreateUpdate.CreateAppointment))
                {
                    telemetryClient?.TrackEvent($"{this.userName} EventGraphHelper OnPrem Event:UpsertEWSAppointment: Creating appointment for:{teamsEvent.Subject}");
                    appointment = new Appointment(service);
                }
                else
                {
                    telemetryClient?.TrackEvent($"{this.userName} EventGraphHelper OnPrem Event:UpsertEWSAppointment: Updating appointment for:{teamsEvent.Subject}");

                    appointment = Appointment.Bind(service, eventId);
                }

                telemetryClient?.TrackEvent($"{this.userName} EventGraphHelper OnPrem Event:UpsertEWSAppointment: Adding appointment contents for:{teamsEvent.Subject}");
                appointment.Subject = teamsEvent.Subject;
                appointment.Body = teamsEvent.Body.Content;
                appointment.Body.BodyType = Exchange.WebServices.Data.BodyType.HTML;
                appointment.Start = DateTime.Parse(teamsEvent.Start.DateTime, CultureInfo.InvariantCulture);
                appointment.End = DateTime.Parse(teamsEvent.End.DateTime, CultureInfo.InvariantCulture);
                appointment.Location = teamsEvent.Location != null ? teamsEvent.Location.DisplayName : string.Empty;

                telemetryClient?.TrackEvent($"{this.userName} EventGraphHelper OnPrem Event:UpsertEWSAppointment: Adding attendies for:{teamsEvent.Subject}");

                if (teamsEvent.Attendees.Any())
                {
                    foreach (var attendee in teamsEvent.Attendees)
                    {
                        if (attendee.Type == 0)
                        {
                            appointment.RequiredAttendees.Add(attendee.EmailAddress.Address);
                        }
                        else
                        {
                            appointment.OptionalAttendees.Add(attendee.EmailAddress.Address);
                        }
                    }
                }

                appointment.ReminderDueBy = DateTime.Now;

                appointment.Recurrence = null;

                if (teamsEvent.Recurrence != null)
                {
                    if (teamsEvent.Recurrence.Range.NumberOfOccurrences > 1)
                    {
                        Recurrence recurrence = new Recurrence.DailyPattern();
                        recurrence.StartDate = appointment.Start.Date;
                        recurrence.NumberOfOccurrences = teamsEvent.Recurrence.Range.NumberOfOccurrences;
                        appointment.Recurrence = recurrence;
                    }
                }

                if (createUpdate.Equals(CreateUpdate.CreateAppointment))
                {
                    appointment.Save(SendInvitationsMode.SendToAllAndSaveCopy);
                    telemetryClient.TrackTrace($"{this.userName} EventGraphHelper OnPrem Event:UpsertEWSAppointment: Event creation SUCCESS");
                    return appointment;
                }
                else
                {
                    SendInvitationsOrCancellationsMode mode = appointment.IsMeeting ?
                        SendInvitationsOrCancellationsMode.SendToAllAndSaveCopy : SendInvitationsOrCancellationsMode.SendToNone;

                    appointment.Update(ConflictResolutionMode.AlwaysOverwrite);
                    telemetryClient.TrackTrace($"{this.userName} EventGraphHelper OnPrem Event:UpsertEWSAppointment: Event update SUCCESS");
                    return appointment;
                }
            }
            catch (Exception ex)
            {
                var ewsExcption = new Exception($"{this.userName} EventGraphHelper OnPrem Event:UpsertEWSAppointment: FAILED with Message:{ex.Message} and Exception:{ex.StackTrace}");
                telemetryClient.TrackException(ewsExcption);
                throw new Exception(ewsExcption.Message);
            }
        }
    }
}
