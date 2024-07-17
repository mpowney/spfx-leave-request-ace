import { ISPFxAdaptiveCard, BaseAdaptiveCardView, IActionArguments } from "@microsoft/sp-adaptive-card-extension-base";
import { Logger, LogLevel } from "@pnp/logging";
import * as strings from "LeaveRequestAdaptiveCardExtensionStrings";
import { LeaveRequest } from "../../../dal/leave/LeaveRequest";
import { LeaveType } from "../../../model/LeaveType";
import { ILeaveRequestAdaptiveCardExtensionProps, ILeaveRequestAdaptiveCardExtensionState } from "../LeaveRequestAdaptiveCardExtension";
import * as card from "./template/QuickViewTemplate.json";

export interface IQuickViewData {
	leaveBalances: [{ leaveTypeLabel: string; leaveBalance: string }];
	leaveCalculatedHours: number;
	leaveStartDate: string;
	leaveFinishDate: string;
	submitting: boolean;
	strings: {};
}

export class QuickView extends BaseAdaptiveCardView<
	ILeaveRequestAdaptiveCardExtensionProps,
	ILeaveRequestAdaptiveCardExtensionState,
	IQuickViewData
> {
	private static numberOfWeekdays(startDate: Date, finishDate: Date): number {
		const diff = (finishDate.getTime() - startDate.getTime()) / 3600000 / 24;
		let workingDate = startDate.getTime();
		let daysToRemove = 0;
		do {
			const today = new Date(workingDate).getDay();
			if (today === 0 || today === 6) daysToRemove++;
			workingDate = workingDate + 3600000 * 24;
		} while (workingDate < finishDate.getTime());
		return diff - daysToRemove;
	}
	public get data(): IQuickViewData {
		Logger.log({ level: LogLevel.Verbose, message: "QuickView data() getter invoked", data: this.state });

		return {
			leaveBalances: [{ leaveTypeLabel: strings.CommonAnnualLeave, leaveBalance: `${this.state.annualLeaveBalance} ${strings.CommonHours}` }],
			leaveCalculatedHours: this.state.leaveCalculatedHours || 0,
			leaveStartDate: `${this.state.leaveStartDate}`,
			leaveFinishDate: `${this.state.leaveFinishDate}`,
			submitting: this.state.submitting,
			strings: strings,
		};
	}

	public get template(): ISPFxAdaptiveCard {
		Logger.log({ level: LogLevel.Verbose, message: "QuickView template getter invoked", data: card });
		return card;
	}

	public async onAction(action: any | IActionArguments): Promise<void> {
		Logger.log({ level: LogLevel.Verbose, message: "QuickView onAction() invoked", data: action });
		if (action.data?.id === "calculateHours") {
			const startDate = new Date(action.data.leaveStartDate);
			const finishDate = new Date(action.data.leaveFinishDate);
			const hours = QuickView.numberOfWeekdays(startDate, finishDate) * 7.8;
			this.setState({
				leaveCalculatedHours: hours,
				leaveStartDate: action.data.leaveStartDate,
				leaveFinishDate: action.data.leaveFinishDate,
			});
		}
		if (action.data?.id === "submitForApproval") {
			this.setState({
				submitting: true,
			});

			const startDate = new Date(action.data.leaveStartDate);
			const finishDate = new Date(action.data.leaveFinishDate);

			const request = new LeaveRequest(this.context);
			const id = await request.submitLeaveRequest(
				this.context.pageContext.user.loginName,
				LeaveType.Annual,
				startDate,
				finishDate,
				action.data.leaveCalculatedHours
			);

			if (id !== 0) {
				this.setState({
					submitting: false,
					submitted: true,
					unnapprovedAnnualLeaveBalance: Number(this.state.unnapprovedAnnualLeaveBalance || 0) + Number(action.data.leaveCalculatedHours),
					leaveStartDate: undefined,
					leaveFinishDate: undefined,
					leaveCalculatedHours: undefined
				});
				this.quickViewNavigator.close();
			}
		}
	}
}
