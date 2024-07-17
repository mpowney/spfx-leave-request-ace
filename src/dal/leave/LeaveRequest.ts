import { spfi, SPFx, SPFI, ISPFXContext } from "@pnp/sp";
import { Logger, LogLevel } from "@pnp/logging";
import { AdaptiveCardExtensionContext } from "@microsoft/sp-adaptive-card-extension-base";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/fields";
import { Constants } from "../../common/Constants";
import { IList, IListInfo } from "@pnp/sp/lists";
import { LeaveType } from "../../model/LeaveType";
import { ChoiceFieldFormatType } from "@pnp/sp/fields";
import { ISharePointList } from "./ISharePointList";
import { ILeaveRequest } from "../../model/ILeaveRequest";

export class LeaveRequest implements ISharePointList {
	private sp: SPFI;
	public list?: IList;

	constructor(protected context: ISPFXContext | AdaptiveCardExtensionContext) {
		this.sp = spfi().using(SPFx(context as any));
	}

	public async checkProvisioned(): Promise<boolean> {
		try {
			if (this.list) {
				return true;
			}
			const list = await this.sp.web.lists.getByTitle(Constants.LeaveRequestListName)();
			if (list) {
				this.list = this.sp.web.lists.getByTitle(Constants.LeaveRequestListName);
			}
			return true;
		} catch (ex) {
			Logger.log({
				level: LogLevel.Error,
				message: `Error occurred in LeaveRequest checkForProvisioned with list title ${Constants.LeaveRequestListName}`,
				data: ex,
			});
		}
		return false;
	}

	public async ensureProvisioned(): Promise<boolean> {
		if (!(await this.checkProvisioned())) {
			Logger.log({ level: LogLevel.Info, message: `LeaveRequest ensureProvisioned - provisioning list ${Constants.LeaveRequestListName}` });

			try {
				const listInfo: Partial<IListInfo> = {
					OnQuickLaunch: false,
				};
				await this.sp.web.lists.add(
					Constants.LeaveRequestListName,
					Constants.LeaveRequestListDescription,
					100,
					false,
					listInfo
				);
				if (await this.checkProvisioned()) {
					await this.list?.fields.addChoice(Constants.LeaveRequestFieldType, {
						EditFormat: ChoiceFieldFormatType.Dropdown,
						FillInChoice: false,
						Choices: [
							LeaveType.Annual,
							LeaveType.Sick,
							LeaveType.Parental,
							LeaveType.Family,
							LeaveType.CommunityService,
							LeaveType.LongService,
						],
					});
					await this.list?.fields.addText(Constants.LeaveRequestFieldUser);
					await this.list?.fields.addDateTime(Constants.LeaveRequestFieldStart);
					await this.list?.fields.addDateTime(Constants.LeaveRequestFieldFinish);
					await this.list?.fields.addNumber(Constants.LeaveRequestFieldCalculatedAmount);
					await this.list?.fields.addBoolean(Constants.LeaveRequestFieldApproved, { Required: false });
				}
			} catch (ex) {
				Logger.log({
					level: LogLevel.Error,
					message: `Error occurred in LeaveRequest ensureProvisioned with list title ${Constants.LeaveRequestListName}`,
					data: ex,
				});
			}
		}
		return await this.checkProvisioned();
	}

	public async ensureSampleData(): Promise<boolean> {
		return true;
	}

	public async submitLeaveRequest(username: string, type: LeaveType, startDate: Date, finishDate: Date, calculatedHours: number): Promise<number> {
		if (!(await this.checkProvisioned())) {
			Logger.log({ level: LogLevel.Warning, message: `LeaveRequest submitLeaveRequest - list not provisioned` });
			await this.ensureProvisioned();
		}
		const response = await this.list?.items.add({
			Title: `${username}-${type}-${startDate.toISOString()}`,
			[Constants.LeaveRequestFieldUser]: username,
			[Constants.LeaveRequestFieldType]: type,
			[Constants.LeaveRequestFieldStart]: startDate.toISOString(),
			[Constants.LeaveRequestFieldFinish]: finishDate.toISOString(),
			[Constants.LeaveRequestFieldCalculatedAmount]: calculatedHours,
			[Constants.LeaveRequestFieldApproved]: false,
		});
		if (response?.data) {
			return response.data?.Id;
		} else {
			Logger.log({ level: LogLevel.Error, message: `LeaveRequest submitLeaveRequest - failed to create request record` });
			return 0;
		}
	}

	private static mapSPItem = (spListItem: any): ILeaveRequest => {
		return {
			username: spListItem[Constants.LeaveRequestFieldUser],
			type: spListItem[Constants.LeaveRequestFieldType],
			startDate: spListItem[Constants.LeaveRequestFieldStart],
			finishDate: spListItem[Constants.LeaveRequestFieldFinish],
			calculatedHours: spListItem[Constants.LeaveRequestFieldCalculatedAmount],
			approved: spListItem[Constants.LeaveRequestFieldApproved],
		};
	};

	public async getLeaveRequests(username: string, type: LeaveType, approved: boolean): Promise<ILeaveRequest[]> {
		if (!(await this.checkProvisioned())) {
			Logger.log({ level: LogLevel.Warning, message: `LeaveRequest getLeaveRequests - list not provisioned` });
			await this.ensureProvisioned();
		}
		try {
			const items = await this.list?.items
				.filter(`${Constants.LeaveRequestFieldUser} eq '${username}' and ${Constants.LeaveRequestFieldType} eq '${type}' and ${Constants.LeaveRequestFieldApproved} eq ${approved ? 1 : 0}`)
				.select(
					Constants.LeaveRequestFieldUser,
					Constants.LeaveRequestFieldType,
					Constants.LeaveRequestFieldStart,
					Constants.LeaveRequestFieldFinish,
					Constants.LeaveRequestFieldCalculatedAmount,
					Constants.LeaveRequestFieldApproved
				)();
			if (items?.length === 0) {
				Logger.log({ level: LogLevel.Info, message: `LeaveRequest getLeaveRequests - no records for user ${username}` });
			} else {
				return items?.map(LeaveRequest.mapSPItem) || [];
			}
		} catch (ex) {
			Logger.log({
				level: LogLevel.Error,
				message: `Error occurred in LeaveRequest getLeaveRequests with username ${username}, type ${type}, and approved ${approved}`,
				data: ex,
			});
		}
        return [];
	}
}
