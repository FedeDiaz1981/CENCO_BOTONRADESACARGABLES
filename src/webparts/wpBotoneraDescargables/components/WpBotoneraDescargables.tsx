import * as React from 'react';
import styles from './WpBotoneraDescargables.module.scss';
import type { IWpBotoneraDescargablesProps } from './IWpBotoneraDescargablesProps';
import { SPHttpClient } from '@microsoft/sp-http';

// Fluent UI
import {
  Stack,
  DefaultButton,
  Spinner,
  SpinnerSize,
  MessageBar,
  MessageBarType,
  Text
} from '@fluentui/react';
import { initializeIcons } from '@fluentui/react/lib/Icons';

interface IState {
  loading: boolean;
  error?: string;
  items: Array<{ id: number; title: string; url: string }>;
}

const LIST_TITLE = 'Descargables';
const TITLE_FIELD = 'Title';
const URL_FIELD = 'Url';

export default class WpBotoneraDescargables extends React.Component<IWpBotoneraDescargablesProps, IState> {
  public state: IState = { loading: true, items: [] };

  public async componentDidMount(): Promise<void> {
    initializeIcons();
    try {
      const items = await this.loadItems();
      this.setState({ items, loading: false });
    } catch (e: any) {
      this.setState({ loading: false, error: e?.message || 'Error al cargar datos' });
    }
  }

  private async loadItems(): Promise<Array<{ id: number; title: string; url: string }>> {
    const { siteUrl, spHttpClient, applyFilter, filterFieldInternalName, filterValue } = this.props;
    const esc = (s: string) => String(s || '').replace(/'/g, "''");

    // SELECT base + el campo de filtro si viene
    const extraField = applyFilter && filterFieldInternalName ? `,${filterFieldInternalName}` : '';
    const select = `$select=Id,${TITLE_FIELD},${URL_FIELD}${extraField}`;

    // FILTER dinámico (solo si se pidió y ambos valores están completos)
    const filter =
      applyFilter && filterFieldInternalName && filterValue
        ? `&$filter=${esc(filterFieldInternalName)} eq '${esc(filterValue)}'`
        : '';

    const orderBy = `&$orderby=Id asc`;
    const url =
      `${siteUrl}/_api/web/lists/getByTitle('${esc(LIST_TITLE)}')/items?${select}${filter}${orderBy}`;

    const res = await spHttpClient.get(url, SPHttpClient.configurations.v1);
    if (!res.ok) throw new Error(`No se pudo obtener la lista (${res.status})`);
    const data = await res.json();

    const rows: any[] = data?.value || [];
    return rows
      .map(r => ({
        id: r.Id,
        title: String(r?.[TITLE_FIELD] ?? ''),
        url: String(r?.[URL_FIELD] ?? '')
      }))
      .filter(x => x.title && x.url);
  }

  // Ícono por título; fallback a click (compat ES5)
  private getIconName(title: string): string {
    const t = (title || '').toLowerCase();
    const has = (s: string) => t.indexOf(s) !== -1;

    if (has('descargar') || has('download')) return 'CloudDownload';
    if (has('material') || has('capacit')) return 'Education';
    if (has('contacto') || has('distrib')) return 'Teamwork';
    if (has('doc') || has('documento')) return 'OpenFile';
    return 'HandPointer';
  }

  public render(): React.ReactElement<IWpBotoneraDescargablesProps> {
    const { loading, error, items } = this.state;

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
            <MessageBar messageBarType={MessageBarType.error} isMultiline={false} className={styles.error_38cc0773}>
              Error: {error}
            </MessageBar>
          </div>
        )}

        {!loading && !error && (
          <Stack
            horizontal
            wrap
            horizontalAlign="center"
            verticalAlign="center"
            tokens={{ childrenGap: 12, padding: 12 }}
            className={styles.links_38cc0773}
          >
            {items.length === 0 && <Text>No hay elementos para mostrar.</Text>}

            {items.map(it => (
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
        )}
      </section>
    );
  }
}
