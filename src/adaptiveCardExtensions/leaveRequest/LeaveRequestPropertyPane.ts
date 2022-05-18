import { IPropertyPaneConfiguration, PropertyPaneTextField } from "@microsoft/sp-property-pane";
import * as strings from "LeaveRequestAdaptiveCardExtensionStrings";

export class LeaveRequestPropertyPane {
	public getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
		return {
			pages: [
				{
					header: { description: strings.PropertyPaneDescription },
					groups: [
						{
							groupFields: [
								PropertyPaneTextField("title", {
									label: strings.TitleFieldLabel,
								}),
							],
						},
					],
				},
			],
		};
	}
}
