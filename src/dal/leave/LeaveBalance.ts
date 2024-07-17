import { spfi, SPFx, SPFI } from "@pnp/sp";
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

export class LeaveBalance implements ISharePointList {
	private sp: SPFI;
	public list?: IList;

	constructor(protected context: AdaptiveCardExtensionContext) {
		this.sp = spfi().using(SPFx(context as any));
	}

	public async checkProvisioned(): Promise<boolean> {
		try {
			if (this.list) {
				return true;
			}
			const list = await this.sp.web.lists.getByTitle(Constants.LeaveBalanceListName)();
			if (list) {
				this.list = this.sp.web.lists.getByTitle(Constants.LeaveBalanceListName);
			}
			return true;
		} catch (ex) {
			Logger.log({
				level: LogLevel.Error,
				message: `Error occurred in LeaveBalance checkForProvisioned with list title ${Constants.LeaveBalanceListName}`,
				data: ex,
			});
		}
		return false;
	}

	public async ensureProvisioned(): Promise<boolean> {
		if (!(await this.checkProvisioned())) {
			Logger.log({ level: LogLevel.Info, message: `LeaveBalance ensureProvisioned - provisioning list ${Constants.LeaveBalanceListName}` });

			try {
				const listInfo: Partial<IListInfo> = {
					OnQuickLaunch: false,
				};
				await this.sp.web.lists.add(
					Constants.LeaveBalanceListName,
					Constants.LeaveBalanceListDescription,
					100,
					false,
					listInfo
				);
				if (await this.checkProvisioned()) {
					await this.list?.fields.addChoice(Constants.LeaveBalanceFieldType, {
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
					await this.list?.fields.addText(Constants.LeaveBalanceFieldUser);
					await this.list?.fields.addNumber(Constants.LeaveBalanceFieldBalance);
				}
				await this.ensureSampleData();
				return true;
		
			} catch (ex) {
				Logger.log({
					level: LogLevel.Error,
					message: `Error occurred in LeaveBalance ensureProvisioned with list title ${Constants.LeaveBalanceListName}`,
					data: ex,
				});
			}
			return false;
		}
		return true;
	}

	public async ensureSampleData(): Promise<boolean> {
		if (await this.checkProvisioned()) {
			try {
				if (this.list && (await this.list.items.top(1)()).length === 0) {
					await this.list?.items.add({
						Title: `${this.context.pageContext.user.loginName}-${LeaveType.Annual}`,
						[Constants.LeaveBalanceFieldUser]: this.context.pageContext.user.loginName,
						[Constants.LeaveBalanceFieldType]: LeaveType.Annual,
						[Constants.LeaveBalanceFieldBalance]: 70.2,
					});
				}
			} catch (ex) {
				Logger.log({
					level: LogLevel.Error,
					message: `Error occurred in LeaveBalance ensureSampleData with list title ${Constants.LeaveBalanceListName}`,
					data: ex,
				});
				return false;
			}
		}
		return false;
	}

	public async getLeaveBalance(username: string, type: LeaveType): Promise<number> {
		if (!(await this.checkProvisioned())) {
			Logger.log({ level: LogLevel.Warning, message: `LeaveBalance getLeaveBalance - list not provisioned` });
			await this.ensureProvisioned();
		} 
		try {
			const items = await this.list?.items
				.filter(`${Constants.LeaveBalanceFieldUser} eq '${username}' and ${Constants.LeaveBalanceFieldType} eq '${type}'`)
				.select(Constants.LeaveBalanceFieldBalance)();
			if (items?.length === 0) {
				Logger.log({ level: LogLevel.Warning, message: `LeaveBalance getLeaveBalance - record for user ${username} doesn't exist` });
			} else {
				return items ? items[0][Constants.LeaveBalanceFieldBalance] : 0;
			}
		} catch (ex) {
			Logger.log({
				level: LogLevel.Error,
				message: `Error occurred in LeaveBalance getLeaveBalance with username ${username} and type ${type}`,
				data: ex,
			});
		}
		return 0;
	}
}
