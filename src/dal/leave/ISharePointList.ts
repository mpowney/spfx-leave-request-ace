import { IList } from "@pnp/sp/lists";

export interface ISharePointList {
	list?: IList;
	checkProvisioned(): Promise<boolean>;
	ensureProvisioned(): Promise<boolean>;
	ensureSampleData(): Promise<boolean>;
}
