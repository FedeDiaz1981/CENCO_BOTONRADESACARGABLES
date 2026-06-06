import * as React from "react";
import styles from "./WpBotoneraDescargables.module.scss";
import type { IWpBotoneraDescargablesProps } from "./IWpBotoneraDescargablesProps";
import { SPHttpClient } from "@microsoft/sp-http";

import {
  Stack,
  DefaultButton,
  Spinner,
  SpinnerSize,
  MessageBar,
  MessageBarType,
  Text,
} from "@fluentui/react";
import { initializeIcons } from "@fluentui/react/lib/Icons";

interface IState {
  loading: boolean;
  error?: string;
  items: Array<{ id: number; title: string; url: string }>;

  // NUEVO: datos de proveedor
  proveedorNombre?: string;
  proveedorRuc?: string;
}

const LIST_TITLE = "Descargables";
const TITLE_FIELD = "Title";
const URL_FIELD = "Url";

// NUEVO: lista de proveedores
const PROV_LIST_TITLE = "Proveedores";
const PROV_TITLE_FIELD = "Title"; // nombre proveedor
const PROV_RUC_FIELD = "RUC"; // campo RUC

export default class WpBotoneraDescargables extends React.Component<
  IWpBotoneraDescargablesProps,
  IState
> {
  public state: IState = { loading: true, items: [] };

  public async componentDidMount(): Promise<void> {
    initializeIcons();

    const doc = window.top ? window.top.document : document;
    const overlay = doc.getElementById("cenco-overlay");
    if (overlay) {
      setTimeout(() => overlay.remove(), 500);
    }

    try {
      const items = await this.loadItems();
      this.setState({ items, loading: false });
    } catch (e: any) {
      this.setState({
        loading: false,
        error: e?.message || "Error al cargar datos",
      });
    }

    // si el toggle de proveedor está activo, cargamos proveedor
    if (this.props.showProveedor) {
      void this.loadProveedor();
    }
  }

  // si cambia el toggle de proveedor en caliente
  public async componentDidUpdate(
    prevProps: IWpBotoneraDescargablesProps
  ): Promise<void> {
    if (!prevProps.showProveedor && this.props.showProveedor) {
      // se activó -> cargamos proveedor
      void this.loadProveedor();
    }

    if (prevProps.showProveedor && !this.props.showProveedor) {
      // se desactivó -> limpiamos labels
      this.setState({ proveedorNombre: undefined, proveedorRuc: undefined });
    }
  }

  private async loadItems(): Promise<
    Array<{ id: number; title: string; url: string }>
  > {
    const {
      siteUrl,
      spHttpClient,
      applyFilter,
      filterFieldInternalName,
      filterValue,
    } = this.props;

    const esc = (s: string) => String(s || "").replace(/'/g, "''");

    const extraField =
      applyFilter && filterFieldInternalName
        ? `,${filterFieldInternalName}`
        : "";
    const select = `$select=Id,${TITLE_FIELD},${URL_FIELD}${extraField}`;

    const filter =
      applyFilter && filterFieldInternalName && filterValue
        ? `&$filter=${esc(filterFieldInternalName)} eq '${esc(filterValue)}'`
        : "";

    const orderBy = `&$orderby=Id asc`;
    const url = `${siteUrl}/_api/web/lists/getByTitle('${esc(
      LIST_TITLE
    )}')/items?${select}${filter}${orderBy}`;

    const res = await spHttpClient.get(url, SPHttpClient.configurations.v1);
    if (!res.ok) {
      throw new Error(`No se pudo obtener la lista (${res.status})`);
    }

    const data = await res.json();
    const rows: any[] = data?.value || [];

    return rows
      .map((r) => ({
        id: r.Id,
        title: String(r?.[TITLE_FIELD] ?? ""),
        url: String(r?.[URL_FIELD] ?? ""),
      }))
      .filter((x) => x.title && x.url);
  }

  private async loadProveedor(): Promise<void> {
    const { siteUrl, spHttpClient, currentUserEmail } = this.props;

    const myEmail = (currentUserEmail || "").toLowerCase();
    if (!myEmail) {
      return;
    }

    const provUrl =
      `${siteUrl}/_api/web/lists/getByTitle('${PROV_LIST_TITLE}')/items` +
      `?$select=${PROV_TITLE_FIELD},${PROV_RUC_FIELD},Usuarios/Id,Usuarios/Title,Usuarios/EMail` +
      `&$expand=Usuarios`;

    const provRes = await spHttpClient.get(
      provUrl,
      SPHttpClient.configurations.v1
    );

    if (!provRes.ok) {
      console.warn(
        "Error HTTP en Proveedores",
        provRes.status,
        provRes.statusText
      );
      return;
    }

    const provData: any = await provRes.json();
    const rows: any[] = provData.value || [];

    let match: any = null;

    for (const row of rows) {
      const usuarios: any[] = row.Usuarios || [];
      for (const u of usuarios) {
        const mail = (u.EMail || "").toLowerCase();
        if (mail && mail === myEmail) {
          match = row;
          break;
        }
      }
      if (match) break;
    }

    if (match) {
      this.setState({
        proveedorNombre: String(match[PROV_TITLE_FIELD] || ""),
        proveedorRuc: String(match[PROV_RUC_FIELD] || ""),
      });
    } else {
      this.setState({
        proveedorNombre: "",
        proveedorRuc: "",
      });
    }
  }

  private getIconName(title: string): string {
    const t = (title || "").toLowerCase();
    const has = (s: string) => t.indexOf(s) !== -1;

    if (has("descargar") || has("download")) return "CloudDownload";
    if (has("material") || has("capacit")) return "Education";
    if (has("contacto") || has("distrib")) return "Teamwork";
    if (has("doc") || has("documento")) return "OpenFile";
    return "HandPointer";
  }

  public render(): React.ReactElement<IWpBotoneraDescargablesProps> {
    const { loading, error, items, proveedorNombre, proveedorRuc } = this.state;
    const { showProveedor } = this.props;

    return (
      <section className={styles.wpBotoneraDescargables_38cc0773}>
        <h2 hidden>Descargables</h2>

        {loading && (
          <div className={styles.wpBotoneraDescargables_38cc0773}>
            <Spinner size={SpinnerSize.large} label="Cargando..." />
          </div>
        )}

        {!loading && error && (
          <div className={styles.wpBotoneraDescargables_38cc0773}>
            <MessageBar
              messageBarType={MessageBarType.error}
              isMultiline={false}
              className={styles.error_38cc0773}
            >
              Error: {error}
            </MessageBar>
          </div>
        )}

        {!loading && !error && (
          <>
            <Stack
              horizontal
              wrap
              horizontalAlign="center"
              verticalAlign="center"
              tokens={{ childrenGap: 12, padding: 12 }}
              className={styles.links_38cc0773}
            >
              {items.length === 0 && (
                <Text>No hay elementos para mostrar.</Text>
              )}

              {items.map((it) => (
                <DefaultButton
                  key={it.id}
                  className={styles.button_38cc0773}
                  text={it.title}
                  iconProps={{ iconName: this.getIconName(it.title) }}
                  href={it.url}
                  target="_blank"
                  rel="noopener noreferrer"
                  allowDisabledFocus
                />
              ))}
            </Stack>

            {showProveedor && (
              <div className={styles.proveedorInfo_38cc0773}>
                <Text>
                  <strong>Proveedor:</strong>{" "}
                  {proveedorNombre && proveedorNombre.trim()
                    ? proveedorNombre
                    : "-"}
                </Text>
                <Text>
                  <strong>RUC:</strong>{" "}
                  {proveedorRuc && proveedorRuc.trim() ? proveedorRuc : "-"}
                </Text>
              </div>
            )}
          </>
        )}
      </section>
    );
  }
}
