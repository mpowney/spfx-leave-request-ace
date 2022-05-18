import {
	BaseImageCardView,
	IImageCardParameters,
	IExternalLinkCardAction,
	IQuickViewCardAction,
	ICardButton,
} from "@microsoft/sp-adaptive-card-extension-base";
import * as strings from "LeaveRequestAdaptiveCardExtensionStrings";
import {
	ILeaveRequestAdaptiveCardExtensionProps,
	ILeaveRequestAdaptiveCardExtensionState,
	QUICK_VIEW_REGISTRY_ID,
} from "../LeaveRequestAdaptiveCardExtension";

export class CardView extends BaseImageCardView<ILeaveRequestAdaptiveCardExtensionProps, ILeaveRequestAdaptiveCardExtensionState> {
	/**
	 * Buttons will not be visible if card size is 'Medium' with Image Card View.
	 * It will support up to two buttons for 'Large' card size.
	 */
	public get cardButtons(): [ICardButton] | [ICardButton, ICardButton] | undefined {
		return this.state.loading
			? undefined
			: [
					{
						title: strings.CommonRequestLeave,
						action: {
							type: "QuickView",
							parameters: {
								view: QUICK_VIEW_REGISTRY_ID,
							},
						},
					},
			  ];
	}

	public get data(): IImageCardParameters {
		const cardImage: string = require("../assets/suitcase.png");
		const iconUrl: string = "https://raw.githubusercontent.com/microsoft/fluentui-system-icons/master/assets/Brightness%20High/SVG/ic_fluent_brightness_high_24_regular.svg";
		return this.state.loading
			? {
					primaryText: strings.PhraseLookingUpYourAnnualLeave,
					imageUrl: cardImage,
					iconProperty: iconUrl,
					title: strings.CommonLeaveRequest,
			  }
			: this.state.annualLeaveBalance && this.state.annualLeaveBalance > 0
			? {
					primaryText:
						(this.state.unnapprovedAnnualLeaveBalance || 0) > 0
							? strings.PhraseYouHaveHoursAnnualLeaveWithUnapproved.replace("{0}", `${this.state.annualLeaveBalance}`).replace(
									"{1}",
									`${this.state.unnapprovedAnnualLeaveBalance || 0}`
							  )
							: strings.PhraseYouHaveHoursAnnualLeave.replace("{0}", `${this.state.annualLeaveBalance}`),

					imageUrl: cardImage,
					iconProperty:iconUrl,
					title: strings.CommonLeaveRequest,
			  }
			: {
					primaryText: strings.PhraseCannotFindAnyAnnualLeave,
					imageUrl: cardImage,
					iconProperty: iconUrl,
					title: strings.CommonLeaveRequest,
			  };
	}

	public get onCardSelection(): IQuickViewCardAction | IExternalLinkCardAction | undefined {
		return this.state.loading
			? undefined
			: {
					type: "QuickView",
					parameters: {
						view: QUICK_VIEW_REGISTRY_ID,
					},
			  };
	}
}
