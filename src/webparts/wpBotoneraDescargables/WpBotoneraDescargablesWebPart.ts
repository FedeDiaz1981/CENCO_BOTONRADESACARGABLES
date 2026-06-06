import * as React from "react";
import * as ReactDom from "react-dom";
import { Version, Log } from "@microsoft/sp-core-library";
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneToggle,
  PropertyPaneDropdown,
  IPropertyPaneDropdownOption,
} from "@microsoft/sp-webpart-base";
import { SPHttpClient } from "@microsoft/sp-http";
import { initializeIcons } from "@fluentui/react/lib/Icons";

import WpBotoneraDescargables from "./components/WpBotoneraDescargables";
import { IWpBotoneraDescargablesProps } from "./components/IWpBotoneraDescargablesProps";

export interface IWpBotoneraDescargablesWebPartProps {
  applyFilter: boolean;
  filterFieldInternalName: string;
  filterValue: string;

  // nuevas
  ocultarRibbon: boolean;
  grupoRibbon: string; // nombre del grupo seleccionado
  mostrarProveedor: boolean; // <- nombre definitivo
}

const LOG_SOURCE = "WpBotoneraDescargablesWebPart";
const HIDE_STYLE_ID = "cenco-theme-styles";

export default class WpBotoneraDescargablesWebPart extends BaseClientSideWebPart<IWpBotoneraDescargablesWebPartProps> {
  private _siteGroups: IPropertyPaneDropdownOption[] = [];
  private _groupsLoaded: boolean = false;

  public async onInit(): Promise<void> {
    initializeIcons();
    this._injectBaseUi();

    // Header SIEMPRE visible.
    // Solo manejamos la ribbon según la propiedad.
    if (this.properties.ocultarRibbon) {
      this._hideRibbonEverywhere();
    } else {
      this._showRibbonEverywhere();
    }

    // Si hay grupo configurado, siempre evaluamos membresía
    if (this.properties.grupoRibbon) {
      await this._checkMembership(this.properties.grupoRibbon);
    }

    return super.onInit();
  }

  public render(): void {
    const element: React.ReactElement<IWpBotoneraDescargablesProps> =
      React.createElement(WpBotoneraDescargables, {
        siteUrl: this.context.pageContext.site.absoluteUrl,
        spHttpClient: this.context.spHttpClient,
        applyFilter: this.properties.applyFilter,
        filterFieldInternalName: this.properties.filterFieldInternalName,
        filterValue: this.properties.filterValue,
        showProveedor: this.properties.mostrarProveedor,
        currentUserLogin: this.context.pageContext.user.loginName,
        currentUserEmail: this.context.pageContext.user.email,
      });

    ReactDom.render(element, this.domElement);
  }

  // =========================
  //  UI base (header + css)
  // =========================
  private _injectBaseUi(): void {
    const doc = window.top ? window.top.document : document;

    if (!doc.getElementById(HIDE_STYLE_ID)) {
      const style = doc.createElement("style");
      style.id = HIDE_STYLE_ID;
      style.innerHTML = `
        :root {
          --cenco-primary: #005596;
          --cenco-primary-light: #0072BC;
          --cenco-accent: #F5A623;
          --cenco-bg: #F5F5F5;
          --cenco-text: #333333;
          --cenco-border: #D9D9D9;
        }

        body {
          background: var(--cenco-bg) !important;
          margin: 0 !important;
          padding: 0 !important;
          font-family: 'Segoe UI Variable','Segoe UI',Inter,Roboto,Arial,sans-serif !important;
          color: var(--cenco-text) !important;
        }

        /* compensar header fijo para todo el contenido */
        #s4-bodyContainer,
        #contentBox,
        #spPageChrome,
        .SPCanvas-canvas,
        .CanvasComponent {
          margin: 96px auto 40px auto !important;
          max-width: 1100px !important;
        }

        /* HEADER CENCOSUD */
        #cenco-header {
          position: fixed !important;
          top: 0;
          left: 0;
          right: 0;
          height: 100px;
          display: flex;
          align-items: center;
          justify-content: space-between;
          padding: 0 32px;
          box-sizing: border-box;
          background: linear-gradient(90deg, var(--cenco-primary) 0%, var(--cenco-primary-light) 60%) !important;
          color: #ffffff !important;
          font-size: 18px;
          font-weight: 600;
          box-shadow: 0 4px 14px rgba(0,0,0,.12);
          z-index: 20000;
        }

        #cenco-header-left {
          display: flex;
          align-items: center;
          gap: 12px;
          min-width: 0;
        }

        #cenco-header-title {
          white-space: nowrap;
          overflow: hidden;
          text-overflow: ellipsis;
          font-size: 1.5rem;
        }

        #cenco-header-flag {
          border-radius: 40px;
          overflow: hidden;
        }

        #cenco-header-center {
          flex: 1 1 auto;
          max-width: 560px;
          padding: 0 32px;
          box-sizing: border-box;
        }

        #cenco-header-search {
          width: 100%;
          height: 36px;
          border-radius: 999px;
          border: none;
          padding: 0 14px;
          font-size: 13px;
          color: var(--cenco-text);
          box-shadow: 0 0 0 1px rgba(255,255,255,0.5);
          background: rgba(255,255,255,0.95);
          outline: none;
          display:none !important;
        }

        #cenco-header-search::placeholder {
          color: rgba(0,0,0,0.45);
        }

        #cenco-header-right {
          display: flex;
          align-items: center;
          gap: 12px;
          font-size: 13px;
          font-weight: 400;
          opacity: 0.9;
        }

        #cenco-header-pill {
          padding: 4px 10px;
          border-radius: 999px;
          background: rgba(0,0,0,0.08);
          border: 1px solid rgba(255,255,255,0.3);
          font-size: 12px;
        }

        @media (max-width: 900px) {
          #cenco-header {
            padding: 0 16px;
          }
          #cenco-header-center {
            display: none;
          }
        }
      `;
      doc.head.appendChild(style);
    }

    if (!doc.getElementById("cenco-header")) {
      const header = doc.createElement("div");
      header.id = "cenco-header";
      header.innerHTML = `
        <div id="cenco-header-left">
          <div id="cenco-header-flag">
            <svg xmlns="http://www.w3.org/2000/svg" width="30" height="20" viewBox="0 0 3 2">
              <rect width="1" height="2" x="0" fill="#D91023"></rect>
              <rect width="1" height="2" x="2" fill="#D91023"></rect>
              <rect width="1" height="2" x="1" fill="#ffffff"></rect>
            </svg>
          </div>
          <div id="cenco-header-title">Portal de Documentación de Transportes</div>
        </div>

        <div id="cenco-header-center">
          <input
            id="cenco-header-search"
            type="search"
            placeholder="Buscar documentación, normas, instructivos..."
          />
        </div>

        <div id="cenco-header-right">
          <span id="cenco-header-pill">Transporte</span>
          <span>Perú</span>
          <button id="cenco-toggle-header"
            style="
              display:none;
              margin-left:8px;
              padding:4px 10px;
              border-radius:999px;
              border:1px solid rgba(255,255,255,0.7);
              background:rgba(0,0,0,0.15);
              color:#fff;
              font-size:12px;
              cursor:pointer;
            ">
            Mostrar edición
          </button>
        </div>
      `;
      doc.body.prepend(header);
    }
  }

  // =========================
  //  Ocultar / mostrar ribbon
  // =========================
  private _hideRibbonEverywhere(): void {
    const doc = window.top ? window.top.document : document;

    const selectors = [
      // Clásico
      "#SuiteNavWrapper",
      "#suiteBar",
      "#s4-ribbonrow",
      "#s4-ribboncont",
      "#sideNavBox",
      "#contentBox > #sideNavBox",
      "#DeltaTopNavigation",
      "#DeltaPlaceHolderLeftNavBar",
      ".ms-breadcrumb-box",
      ".ms-breadcrumb-dropdownBox",

      // Modern header + command bar
      'div[data-automationid="SiteHeader"]',
      'div[data-automationid="SiteHeader-host"]',
      'div[data-automationid="PageHeader"]',
      'div[data-automationid="PageCommandBar"]',
      'div[data-automationid="CommandBar"]',

      // Menús laterales / navegación
      'div[data-automationid="QuickLaunch"]',
      'div[data-automationid="VerticalNav"]',
      'nav[aria-label="Navegación de sitio"]',
      'nav[aria-label="Site"]',
      "#spLeftNav",
      ".sp-appBar",       // barra vertical izquierda moderna
      ".ms-FlexPane"      // panel lateral (a veces para nav)
    ];

    selectors.forEach((sel) => {
      const el = doc.querySelector(sel) as HTMLElement | null;
      if (el) {
        el.style.setProperty("display", "none", "important");
        el.style.setProperty("visibility", "hidden", "important");
        el.style.setProperty("height", "0", "important");
        el.style.setProperty("width", "0", "important");
      }
    });
  }

  private _showRibbonEverywhere(): void {
    const doc = window.top ? window.top.document : document;

    const selectors = [
      // Clásico
      "#SuiteNavWrapper",
      "#suiteBar",
      "#s4-ribbonrow",
      "#s4-ribboncont",
      "#sideNavBox",
      "#contentBox > #sideNavBox",
      "#DeltaTopNavigation",
      "#DeltaPlaceHolderLeftNavBar",
      ".ms-breadcrumb-box",
      ".ms-breadcrumb-dropdownBox",

      // Modern header + command bar
      'div[data-automationid="SiteHeader"]',
      'div[data-automationid="SiteHeader-host"]',
      'div[data-automationid="PageHeader"]',
      'div[data-automationid="PageCommandBar"]',
      'div[data-automationid="CommandBar"]',

      // Menús laterales / navegación
      'div[data-automationid="QuickLaunch"]',
      'div[data-automationid="VerticalNav"]',
      'nav[aria-label="Navegación de sitio"]',
      'nav[aria-label="Site"]',
      "#spLeftNav",
      ".sp-appBar",
      ".ms-FlexPane"
    ];

    selectors.forEach((sel) => {
      const el = doc.querySelector(sel) as HTMLElement | null;
      if (el) {
        el.style.removeProperty("display");
        el.style.removeProperty("visibility");
        el.style.removeProperty("height");
        el.style.removeProperty("width");
      }
    });
  }

  // =========================
  //  Botón de edición en header
  // =========================
  private _enableEditToggle(): void {
    const doc = window.top ? window.top.document : document;
    const btn = doc.getElementById(
      "cenco-toggle-header"
    ) as HTMLButtonElement | null;
    const header = doc.getElementById("cenco-header") as HTMLElement | null;

    if (!btn || !header) {
      return;
    }

    // mostrar botón para miembros del grupo
    btn.style.display = "inline-flex";

    btn.onclick = () => {
      // ocultar header custom
      header.style.display = "none";

      // mostrar ribbon y barra estándar
      this._showRibbonEverywhere();

      const banner = doc.getElementById(
        "cenco-edit-banner"
      ) as HTMLElement | null;
      if (banner) {
        banner.style.display = "none";
      }
    };
  }

  // =========================
  //  Normalizar login
  // =========================
  private _normalizeLogin(login: string): string {
    if (!login) {
      return "";
    }
    const lowered = login.toLowerCase();
    const parts = lowered.split("|");
    return parts[parts.length - 1]; // "i:0#.f|membership|user@x" -> "user@x"
  }

  // =========================
  //  Obtener usuario actual
  // =========================
  private async _getCurrentUserLogin(): Promise<string | null> {
    // 1) Intentar con pageContext
    const u = this.context.pageContext.user;
    if (u && (u.loginName || u.email)) {
      const raw = u.loginName || u.email || "";
      const fromCtx = this._normalizeLogin(raw);
      console.log(
        "Usuario desde pageContext (raw):",
        raw,
        "normalizado:",
        fromCtx
      );
      return fromCtx;
    }

    const webUrl = this.context.pageContext.web.absoluteUrl.replace(/\/$/, "");

    try {
      // 2) Fallback: /currentuser
      let res = await this.context.spHttpClient.get(
        `${webUrl}/_api/web/currentuser`,
        SPHttpClient.configurations.v1,
        {
          headers: {
            Accept: "application/json;odata=verbose",
            "odata-version": "",
          },
        }
      );

      if (res.ok) {
        const data: any = await res.json();
        const login = data?.d?.LoginName || "";
        if (login) {
          const norm = this._normalizeLogin(login);
          console.log("Usuario desde /currentuser:", login, "norm:", norm);
          return norm;
        }
      }

      // 3) Fallback extra: PeopleManager
      res = await this.context.spHttpClient.get(
        `${webUrl}/_api/SP.UserProfiles.PeopleManager/GetMyProperties`,
        SPHttpClient.configurations.v1,
        {
          headers: {
            Accept: "application/json;odata=verbose",
            "odata-version": "",
          },
        }
      );

      if (res.ok) {
        const profile: any = await res.json();
        const login = profile?.d?.AccountName || profile?.d?.Email || "";
        if (login) {
          const norm = this._normalizeLogin(login);
          console.log(
            "Usuario desde GetMyProperties:",
            login,
            "norm:",
            norm
          );
          return norm;
        }
      }

      console.warn("No se pudo obtener usuario mediante ninguna API");
      return null;
    } catch (error) {
      console.error("Error recuperando usuario:", error);
      return null;
    }
  }

  // =========================
  //  Chequear grupo elegido
  // =========================
  private async _checkMembership(groupName: string): Promise<void> {
    try {
      const webUrl = this.context.pageContext.web.absoluteUrl.replace(
        /\/$/,
        ""
      );

      const myLogin = await this._getCurrentUserLogin();
      if (!myLogin) {
        console.warn(
          "No se pudo identificar usuario actual, no se mostrará botón."
        );
        return;
      }

      const grpRes = await this.context.spHttpClient.get(
        `${webUrl}/_api/web/sitegroups/getbyname('${groupName}')/users?$select=LoginName`,
        SPHttpClient.configurations.v1,
        {
          headers: {
            Accept: "application/json;odata=verbose",
            "odata-version": "",
          },
        }
      );
      const grpData: any = await grpRes.json();
      const users: Array<{ LoginName: string }> =
        grpData.d?.results || grpData.value || [];

      const normalizedMyLogin = this._normalizeLogin(myLogin);
      const normalizedUsers = users.map((u) =>
        this._normalizeLogin(u.LoginName || "")
      );

      const isMember = normalizedUsers.some(
        (ln) => ln === normalizedMyLogin
      );

      console.log("grupoRibbon:", groupName);
      console.log("current user (norm):", normalizedMyLogin);
      console.log("users norm:", normalizedUsers);
      console.log("isMember:", isMember);

      if (isMember) {
        this._addEditBanner(groupName);
        this._enableEditToggle();
      }
      // Si no es miembro, no hay botón. La ribbon queda según ocultarRibbon.
    } catch (err) {
      Log.warn(
        LOG_SOURCE,
        "Error verificando grupo seleccionado: " + String(err)
      );
    }
  }

  private _addEditBanner(groupName: string): void {
    const doc = window.top ? window.top.document : document;
    if (!doc.getElementById("cenco-edit-banner")) {
      const banner = doc.createElement("div");
      banner.id = "cenco-edit-banner";
      banner.innerHTML = `🔧 Modo edición (${groupName})`;
      banner.setAttribute(
        "style",
        `
          position: fixed;
          top: 72px;
          left: 0;
          width: 100%;
          background: #e0e0e0;
          color: #333;
          font-size: 13px;
          font-weight: 500;
          text-align: center;
          padding: 4px 0;
          z-index: 20001;
          font-family: 'Segoe UI Variable','Segoe UI',Inter,Roboto,Arial,sans-serif;
        `
      );
      doc.body.prepend(banner);
    }
  }

  // =========================
  //  Property Pane
  // =========================
  protected onPropertyPaneConfigurationStart(): void {
    if (!this._groupsLoaded) {
      this._groupsLoaded = true;
      this._loadSiteGroups()
        .then(() => {
          this.context.propertyPane.refresh();
        })
        .catch((err) => {
          Log.warn(LOG_SOURCE, "Error cargando grupos: " + String(err));
        });
    }
  }

  private async _loadSiteGroups(): Promise<void> {
    const siteUrl = this.context.pageContext.site.absoluteUrl.replace(
      /\/$/,
      ""
    );
    const gruposUrl = `${siteUrl}/_api/web/sitegroups?$select=Title&$orderby=Title`;

    const res = await this.context.spHttpClient.get(
      gruposUrl,
      SPHttpClient.configurations.v1,
      {
        headers: {
          Accept: "application/json;odata=verbose",
          "odata-version": "",
        },
      }
    );

    const data: any = await res.json();
    const groups: Array<{ Title: string }> = data.d?.results || [];

    this._siteGroups = groups.map((g) => ({
      key: g.Title,
      text: g.Title,
    }));
    this._siteGroups.unshift({ key: "", text: "(sin grupo)" });
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: { description: "Configuración" },
          groups: [
            {
              groupName: "Filtro",
              groupFields: [
                PropertyPaneToggle("applyFilter", {
                  label: "Aplicar filtro",
                  onText: "Sí",
                  offText: "No",
                }),
                PropertyPaneTextField("filterFieldInternalName", {
                  label: "Campo de filtro (InternalName)",
                }),
                PropertyPaneTextField("filterValue", {
                  label: "Valor a filtrar",
                }),
              ],
            },
            {
              groupName: "Visibilidad de SharePoint",
              groupFields: [
                PropertyPaneToggle("ocultarRibbon", {
                  label: "Ocultar ribbon y navegación",
                  onText: "Ocultar",
                  offText: "No ocultar",
                }),
                PropertyPaneDropdown("grupoRibbon", {
                  label: "Mostrar botón de edición solo para el grupo",
                  options: this._siteGroups,
                  disabled: false,
                }),
              ],
            },
            {
              groupName: "Proveedor",
              groupFields: [
                PropertyPaneToggle("mostrarProveedor", {
                  label: "Proveedor",
                  onText: "Mostrar",
                  offText: "Ocultar",
                }),
              ],
            },
          ],
        },
      ],
    };
  }

  protected onPropertyPaneFieldChanged(
    propertyPath: string,
    oldValue: any,
    newValue: any
  ): void {
    if (propertyPath === "ocultarRibbon") {
      this.properties.ocultarRibbon = !!newValue;

      if (this.properties.ocultarRibbon) {
        this._hideRibbonEverywhere();
      } else {
        this._showRibbonEverywhere();
      }
      this.context.propertyPane.refresh();
    }

    if (propertyPath === "grupoRibbon") {
      this.properties.grupoRibbon = String(newValue || "");

      if (this.properties.grupoRibbon) {
        void this._checkMembership(this.properties.grupoRibbon);
      } else {
        // si se borra el grupo, simplemente no hay botón.
        const doc = window.top ? window.top.document : document;
        const btn = doc.getElementById(
          "cenco-toggle-header"
        ) as HTMLElement | null;
        const banner = doc.getElementById(
          "cenco-edit-banner"
        ) as HTMLElement | null;
        if (btn) btn.style.display = "none";
        if (banner) banner.style.display = "none";
      }
    }

    if (propertyPath === "mostrarProveedor") {
      this.properties.mostrarProveedor = !!newValue;
    }

    super.onPropertyPaneFieldChanged(propertyPath, oldValue, newValue);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse("1.0");
  }
}
