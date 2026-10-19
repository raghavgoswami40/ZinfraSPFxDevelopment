import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneSlider,
  PropertyPaneTextField,
  PropertyPaneToggle,
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';

import PageBanner, { IPageBannerProps } from './components/PageBanner';
import { isValidHex } from '../processPanel/components/roleColors';

export interface IPageBannerWebPartProps {
  title: string;
  height: number;
  darkBackgroundColor: string;
  showSparkles: boolean;
  followPageTheme: boolean;
}

const DEFAULT_BACKGROUND = '#000000';

export default class PageBannerWebPart extends BaseClientSideWebPart<IPageBannerWebPartProps> {

  public render(): void {
    const p = this.properties;
    const element: React.ReactElement<IPageBannerProps> = React.createElement(PageBanner, {
      title: p.title || '',
      height: p.height || 160,
      // A new property name, so a navy value saved by an earlier build no longer overrides black.
      backgroundColor: isValidHex(p.darkBackgroundColor) ? p.darkBackgroundColor : DEFAULT_BACKGROUND,
      showSparkles: p.showSparkles !== false,
      followPageTheme: p.followPageTheme !== false,
    });
    ReactDom.render(element, this.domElement);
  }

  protected onThemeChanged(_currentTheme: IReadonlyTheme | undefined): void { /* no-op */ }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: { description: 'Page Banner' },
          groups: [
            {
              groupName: 'Banner',
              groupFields: [
                PropertyPaneTextField('title', { label: 'Title' }),
                PropertyPaneSlider('height', { label: 'Height (px)', min: 120, max: 480, step: 10, showValue: true }),
                PropertyPaneTextField('darkBackgroundColor', {
                  label: 'Dark theme background colour',
                  onGetErrorMessage: (v: string) =>
                    !v || isValidHex(v) ? '' : 'Enter a hex colour, for example #000000.',
                }),
                PropertyPaneToggle('showSparkles', { label: 'Show starry animation' }),
                PropertyPaneToggle('followPageTheme', {
                  label: 'Allow theme to change',
                  onText: 'Follows the Process Hierarchy / Master RASCI theme',
                  offText: 'Fixed dark theme',
                }),
              ],
            },
          ],
        },
      ],
    };
  }
}
