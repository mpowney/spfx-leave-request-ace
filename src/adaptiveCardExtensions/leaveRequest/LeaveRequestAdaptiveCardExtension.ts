import { IPropertyPaneConfiguration } from "@microsoft/sp-property-pane";
import { BaseAdaptiveCardExtension } from "@microsoft/sp-adaptive-card-extension-base";
import { CardView } from "./cardView/CardView";
import { QuickView } from "./quickView/QuickView";
import { LeaveRequestPropertyPane } from "./LeaveRequestPropertyPane";
import { LeaveBalance } from "../../dal/leave/LeaveBalance";
import { LeaveType } from "../../model/LeaveType";
import { ConsoleListener, Logger, LogLevel } from "@pnp/logging";
import { LeaveRequest } from "../../dal/leave/LeaveRequest";

export interface ILeaveRequestAdaptiveCardExtensionProps {
	title: string;
}

export interface ILeaveRequestAdaptiveCardExtensionState {
	loading: boolean;
	submitting: boolean;
	submitted: boolean;
	annualLeaveBalance?: number;
	unnapprovedAnnualLeaveBalance?: number;
	leaveStartDate?: string;
	leaveFinishDate?: string;
	leaveCalculatedHours?: number;
}

const CARD_VIEW_REGISTRY_ID: string = "LeaveRequest_CARD_VIEW";
export const QUICK_VIEW_REGISTRY_ID: string = "LeaveRequest_QUICK_VIEW";

export default class LeaveRequestAdaptiveCardExtension extends BaseAdaptiveCardExtension<
	ILeaveRequestAdaptiveCardExtensionProps,
	ILeaveRequestAdaptiveCardExtensionState
> {
	private _deferredPropertyPane: LeaveRequestPropertyPane | undefined;

	public async onInit(): Promise<void> {
		
		Logger.subscribe(ConsoleListener("LeaveRequestAdaptiveCardExtension"));
		Logger.activeLogLevel = LogLevel.Verbose;

		this.state = { loading: true, submitting: false, submitted: false };

		this.cardNavigator.register(CARD_VIEW_REGISTRY_ID, () => new CardView());
		this.quickViewNavigator.register(QUICK_VIEW_REGISTRY_ID, () => new QuickView());

		const balance = new LeaveBalance(this.context);
		this.setState({ annualLeaveBalance: await balance.getLeaveBalance(this.context.pageContext.user.loginName, LeaveType.Annual) });

		const request = new LeaveRequest(this.context);
		const unapprovedRequests = await request.getLeaveRequests(this.context.pageContext.user.loginName, LeaveType.Annual, false);
		if (unapprovedRequests) {
			let unnaprovedBalance = 0;
			unapprovedRequests.forEach(request => { unnaprovedBalance = unnaprovedBalance + request.calculatedHours})
			this.setState({ unnapprovedAnnualLeaveBalance: unnaprovedBalance });
		}

		this.setState({ loading: false });

		return Promise.resolve();
	}

	protected loadPropertyPaneResources(): Promise<void> {
		return import(
			/* webpackChunkName: 'LeaveRequest-property-pane'*/
			"./LeaveRequestPropertyPane"
		).then(component => {
			this._deferredPropertyPane = new component.LeaveRequestPropertyPane();
		});
	}

	protected renderCard(): string | undefined {
		return CARD_VIEW_REGISTRY_ID;
	}

	protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
		// eslint-disable-next-line @typescript-eslint/no-non-null-assertion
		return this._deferredPropertyPane!.getPropertyPaneConfiguration();
	}
}
