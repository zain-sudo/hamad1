// <copyright file="register-remove-event.tsx" company="Microsoft">
// Copyright (c) Microsoft. All rights reserved.
// </copyright>

import * as React from "react";
import * as microsoftTeams from "@microsoft/teams-js";
import moment from "moment";
import { WithTranslation, withTranslation } from "react-i18next";
import { TFunction } from "i18next";
import { IEvent } from "../../models/IEvent";
import { ResponseStatus } from "../../constants/constants";
import { EventOperationType } from "../../models/event-operation-type";
import { EventStatus } from "../../models/event-status";
import { getEventAsync } from "../../api/common-api";
import { getUserProfiles } from "../../api/user-group-api"
import { registerToEventAsync, removeEventAsync } from "../../api/user-events-api";
import EventDetails from "../event-operation-task-module/event-details";
import withContext, { IWithContext } from "../../providers/context-provider";
import { Loader } from "@fluentui/react-northstar";

interface IRegisterRemoveEventProps extends IWithContext, WithTranslation {
}

interface IRegisterRemoveEventState {
    isLoading: boolean,
    isOperationInProgress: boolean,
    eventDetails: IEvent | undefined,
    eventCreatedBy: string,
    isErrorGettingEventDetails: boolean,
    isFailedToRegisterOrRemove: boolean,
    eventOperationType: EventOperationType
}

class RegisterRemoveEvent extends React.Component<IRegisterRemoveEventProps, IRegisterRemoveEventState> {
    readonly localize: TFunction;
    isMobileView: boolean;

    constructor(props) {
        super(props);

        this.localize = this.props.t;
        this.isMobileView = false;

        this.state = {
            isLoading: true,
            isOperationInProgress: false,
            eventDetails: undefined,
            eventCreatedBy: "",
            isErrorGettingEventDetails: false,
            isFailedToRegisterOrRemove: false,
            eventOperationType: EventOperationType.None
        }
    }

    componentDidMount() {
        microsoftTeams.initialize();
        this.getEventDetailsAsync();
    }

    /** Gets event details */
    getEventDetailsAsync = async () => {
        let queryParam = new URLSearchParams(window.location.search);
        let eventId = queryParam.get("eventId") ?? "0";
        let teamId = queryParam.get("teamId") ?? "0";
        this.isMobileView = queryParam.get("isMobileView") === "true" ? true : false;

        let response = await getEventAsync(eventId!, teamId!);

        if (response.status === ResponseStatus.OK && response.data) {
            let eventDetails: IEvent = response.data;
            let eventOperationType: EventOperationType = EventOperationType.None;

            if (eventDetails.status === EventStatus.Active && new Date() < moment.utc(eventDetails.endDate).local().toDate()) {
                if (eventDetails.isLoggedInUserRegistered) {
                    eventOperationType = EventOperationType.Remove;
                }
                else if (!eventDetails.isRegistrationClosed && eventDetails.registeredAttendeesCount < eventDetails.maximumNumberOfParticipants && eventDetails.canLoggedInUserRegister) {
                    eventOperationType = EventOperationType.Register;
                }
            }

            this.setState({isLoading: false, eventDetails, eventOperationType }, () => {
                if (this.state.eventDetails) {
                    this.getUserProfileAsync(this.state.eventDetails.createdBy);
                }
            });
        }
        else {
            this.setState({ isLoading: false, isErrorGettingEventDetails: true });
        }
    }

    /**
     * Get event creator information
     * @param userId The user ID of which information to get
     */
    getUserProfileAsync = async (userId: string) => {
        let user: Array<string> = [userId];
        let response = await getUserProfiles(user);

        if (response.status === ResponseStatus.OK && response.data) {
            let userInfo = response.data[0];
            this.setState({ eventCreatedBy: userInfo.displayName });
        }
    }

    /** Called when click on 'Register' or 'Remove' event */
    onRegisterRemoveEvent = async () => {


        let responseSuccess, responseStatusSuccess, responseDataSuccess;

        let teamId = this.state.eventDetails ? this.state.eventDetails.teamId : "0";
        let eventId = this.state.eventDetails ? this.state.eventDetails.eventId : "0";

        switch (this.state.eventOperationType) {
            case EventOperationType.Register:
                this.setState({ isLoading: true });

                responseSuccess = await registerToEventAsync(teamId, eventId);
                responseStatusSuccess = responseSuccess.status === ResponseStatus.OK;
                responseDataSuccess = responseSuccess.data === true;
                break;

            case EventOperationType.Remove:
                this.setState({ isLoading: true });
                responseSuccess = await removeEventAsync(teamId, eventId);
                responseStatusSuccess = responseSuccess.status === ResponseStatus.OK;
                responseDataSuccess = responseSuccess.data === true;
                break;

            default:
                break;
        }

        if (responseSuccess && responseStatusSuccess && responseDataSuccess) {
            microsoftTeams.tasks.submitTask({ isSuccess: true, type: this.state.eventOperationType });
            this.setState({ isFailedToRegisterOrRemove: false, isOperationInProgress: true });
            this.setState({ isLoading: false });
        }
        else {
            this.setState({ isFailedToRegisterOrRemove: true, isOperationInProgress: false });
            this.setState({ isLoading: false });
        }
    }

    /** Renders component */
    render() {
        return (
            <EventDetails
                dir={this.props.dir}
                eventDetails={this.state.eventDetails}
                eventCreatedByName={this.state.eventCreatedBy}
                eventOperationType={this.state.eventOperationType}
                isLoadingEventDetails={this.state.isLoading}
                isFailedToGetEventDetails={this.state.isErrorGettingEventDetails}
                isOperationInProgress={this.state.isOperationInProgress}
                isOperationFailed={this.state.isFailedToRegisterOrRemove}
                isMobileView={this.isMobileView}
                onPerformOperation={this.onRegisterRemoveEvent}
            />
        );
    }
}

export default withTranslation()(withContext(RegisterRemoveEvent));