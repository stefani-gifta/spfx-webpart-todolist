import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import type { IReadonlyTheme } from '@microsoft/sp-component-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './BannerWebPart.module.scss';

// Bundled placeholder — swap /assets/company-logo.png for the real company logo.
// See png.d.ts for the module declaration that makes this import work.
import companyLogoAsset from './assets/company-logo.png';

const strings = {
  PropertyPaneDescription: 'Banner settings',
  BasicGroupName: 'General',
  CompanyLogoAltFieldLabel: 'Company logo alt text'
};

export interface IBannerWebPartProps {
  companyLogoAlt: string;
}

export default class BannerWebPart extends BaseClientSideWebPart<IBannerWebPartProps> {

  public render(): void {
    const displayName = this.context.pageContext.user.displayName || '';
    const firstName = displayName.trim().split(/\s+/)[0] || displayName;
    const initials = this._getInitials(displayName);
    const greeting = this._getGreeting();
    const dateText = this._getFormattedDate();
    const avatarSrc = this._getUserPhotoUrl();

    const mailUrl = 'https://outlook.office.com/mail/';
    const calendarUrl = 'https://outlook.office.com/calendar/';
    const oneDriveUrl = this._getOneDriveUrl();

    const logoAlt = escape(this.properties.companyLogoAlt || 'Company logo');

    this.domElement.innerHTML = `
      <section class="${styles.banner}">
        <div class="${styles.left}">
          <div class="${styles.leftDecor}" aria-hidden="true"></div>
          <img class="${styles.companyLogo}" src="${companyLogoAsset}" alt="${logoAlt}" />
          <div class="${styles.leftContent}">
            <span class="${styles.eyebrow}">${escape(greeting)}</span>
            <h1 class="${styles.name}">${escape(firstName)}</h1>
            <span class="${styles.dateBadge}">${escape(dateText)}</span>
          </div>
        </div>
        <div class="${styles.right}">
          <div class="${styles.avatar}">
            <img
              class="${styles.avatarImg}"
              data-role="avatar-img"
              src="${avatarSrc}"
              alt=""
            />
            <span class="${styles.avatarInitials}">${escape(initials)}</span>
            <span class="${styles.presenceDot}" aria-hidden="true"></span>
          </div>
          <div class="${styles.userName}">${escape(displayName)}</div>
          <div class="${styles.links}">
            <a class="${styles.linkCard}" href="${mailUrl}" target="_blank" rel="noopener noreferrer">
              <span class="${styles.linkIcon}">${this._icon('mail')}</span>
              <span>Email</span>
            </a>
            <a class="${styles.linkCard}" href="${calendarUrl}" target="_blank" rel="noopener noreferrer">
              <span class="${styles.linkIcon}">${this._icon('calendar')}</span>
              <span>Calendar</span>
            </a>
            <a class="${styles.linkCard}" href="${oneDriveUrl}" target="_blank" rel="noopener noreferrer">
              <span class="${styles.linkIcon}">${this._icon('onedrive')}</span>
              <span>OneDrive</span>
            </a>
          </div>
        </div>
      </section>`;

    // If the tenant/site blocks userphoto.aspx or the user has no photo, fall back to initials.
    const avatarImg = this.domElement.querySelector('[data-role="avatar-img"]') as HTMLImageElement | null;
    if (avatarImg) {
      avatarImg.addEventListener('error', () => {
        avatarImg.style.display = 'none';
      });
    }
  }

  // ---------- content helpers ----------

  private _getGreeting(): string {
    const hour = new Date().getHours();
    if (hour < 12) { return 'Good morning'; }
    if (hour < 17) { return 'Good afternoon'; }
    return 'Good evening';
  }

  private _getFormattedDate(): string {
    const locale = this.context.pageContext.cultureInfo
      ? this.context.pageContext.cultureInfo.currentUICultureName
      : undefined;

    try {
      return new Intl.DateTimeFormat(locale, {
        weekday: 'long',
        day: 'numeric',
        month: 'long',
        year: 'numeric'
      }).format(new Date());
    } catch {
      return new Date().toDateString();
    }
  }

  private _getInitials(displayName: string): string {
    const parts = displayName.trim().split(/\s+/).filter(Boolean);
    if (parts.length === 0) { return ''; }
    if (parts.length === 1) { return parts[0].charAt(0).toUpperCase(); }
    return (parts[0].charAt(0) + parts[parts.length - 1].charAt(0)).toUpperCase();
  }

  private _getUserPhotoUrl(): string {
    const loginName = this.context.pageContext.user.loginName;
    const base = this.context.pageContext.web.absoluteUrl;
    return `${base}/_layouts/15/userphoto.aspx?size=L&accountname=${encodeURIComponent(loginName)}`;
  }

  private _getOneDriveUrl(): string {
    const base = this.context.pageContext.web.absoluteUrl;
    return `${base}/_layouts/15/onedrive.aspx`;
  }

  private _icon(name: 'mail' | 'calendar' | 'onedrive'): string {
    switch (name) {
      case 'mail':
        return `<svg width="16" height="16" viewBox="0 0 20 20" fill="none" xmlns="http://www.w3.org/2000/svg"><rect x="2" y="4" width="16" height="12" rx="2" stroke="currentColor" stroke-width="1.5"/><path d="M3 5.5l7 5 7-5" stroke="currentColor" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round"/></svg>`;
      case 'calendar':
        return `<svg width="16" height="16" viewBox="0 0 20 20" fill="none" xmlns="http://www.w3.org/2000/svg"><rect x="2.5" y="3.5" width="15" height="14" rx="2" stroke="currentColor" stroke-width="1.5"/><path d="M2.5 8h15" stroke="currentColor" stroke-width="1.5"/><path d="M6.5 2v3M13.5 2v3" stroke="currentColor" stroke-width="1.5" stroke-linecap="round"/></svg>`;
      case 'onedrive':
        return `<svg width="16" height="16" viewBox="0 0 20 20" fill="none" xmlns="http://www.w3.org/2000/svg"><path d="M6.5 14.5h8a3 3 0 000-6 4.5 4.5 0 00-8.7-1.6A3.5 3.5 0 006.5 14.5z" stroke="currentColor" stroke-width="1.5" stroke-linejoin="round"/></svg>`;
      default:
        return '';
    }
  }

  // ---------- theme ----------

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) { return; }

    const { palette, semanticColors, fonts } = currentTheme;
    this.domElement.classList.toggle(styles.isDarkTheme, !!currentTheme.isInverted);

    if (palette) {
      this.domElement.style.setProperty('--themeDarker', palette.themeDarker || null);
      this.domElement.style.setProperty('--themeDark', palette.themeDark || null);
      this.domElement.style.setProperty('--themeDarkAlt', palette.themeDarkAlt || null);
      this.domElement.style.setProperty('--themePrimary', palette.themePrimary || null);
      this.domElement.style.setProperty('--white', palette.white || null);
      this.domElement.style.setProperty('--neutralLighterAlt', palette.neutralLighterAlt || null);
      this.domElement.style.setProperty('--neutralLight', palette.neutralLight || null);
      this.domElement.style.setProperty('--neutralPrimary', palette.neutralPrimary || null);
      this.domElement.style.setProperty('--neutralSecondary', palette.neutralSecondary || null);
    }

    if (semanticColors) {
      this.domElement.style.setProperty('--bodyBackground', semanticColors.bodyBackground || null);
      this.domElement.style.setProperty('--bodyText', semanticColors.bodyText || null);
    }

    if (fonts && fonts.medium) {
      this.domElement.style.setProperty('--bodyFontFamily', fonts.medium.fontFamily || null);
    }
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('companyLogoAlt', {
                  label: strings.CompanyLogoAltFieldLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
