import * as React from "react";
import * as ReactDom from "react-dom";
import { Version } from "@microsoft/sp-core-library";
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField,
} from "@microsoft/sp-property-pane";
import { BaseClientSideWebPart } from "@microsoft/sp-webpart-base";

import * as strings from "NmdRssWebPartStrings";
import RSSFeedComponent from "./components/NmdRss";

export interface INmdRssWebPartProps {
  description: string;
  feedUrl: string;
}

export default class NmdRssWebPart extends BaseClientSideWebPart<INmdRssWebPartProps> {
  public render(): void {
    const element: React.ReactElement = React.createElement(RSSFeedComponent, {
      feedUrl: this.properties.feedUrl,
      spHttpClient: this.context.spHttpClient,
      siteUrl: this.context.pageContext.web.absoluteUrl,
    });

    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse("1.0");
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription,
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField("feedUrl", {
                  label: "RSS Feed URL",
                }),
              ],
            },
          ],
        },
      ],
    };
  }
}
