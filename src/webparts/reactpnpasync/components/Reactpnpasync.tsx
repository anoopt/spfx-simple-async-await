import * as React from 'react';
import styles from './Reactpnpasync.module.scss';

import { INewsItem } from "../interfaces";
// import pnp
import { Web } from "sp-pnp-js";
// import SPFx Logging system
import { Log } from "@microsoft/sp-core-library";

import { LogHandler, LogLevel } from '../../../common/LogHandler';
import { IReactpnpasyncProps } from './IReactpnpasyncProps';
import { IReactpnpasyncState } from './IReactpnpasyncState';
import { escape } from '@microsoft/sp-lodash-subset';

const LOG_SOURCE: string = 'ReactPnPAsync';
export default class Reactpnpasync extends React.Component<IReactpnpasyncProps, IReactpnpasyncState> {

  constructor(props: IReactpnpasyncProps){
    super(props);
    this.state = {
      items: [],
      errors: [],
      status: "Ready"
    };
    Log._initialize(new LogHandler((window as any).LOG_LEVEL || LogLevel.Error));
    Log.verbose(LOG_SOURCE, "In constrcutor.");
    this._readItems.bind(this);
  }

  public componentDidMount(): void {
    this.setState({
        items: [],
        errors: [],
        status: "Loading"
      });
    this._readItems("News");
  }

  public render(): React.ReactElement<IReactpnpasyncProps> {

    return (
        <div className={styles.container}>
          <div className={`ms-Grid-row ms-bgColor-themeDark ms-fontColor-white ${styles.row}`}>
            <div className="ms-Grid-col ms-u-lg10 ms-u-xl8 ms-u-xlPush2 ms-u-lgPush1">
              <span className="ms-font-xl ms-fontColor-white">React PnP JS Async Await</span>
              <div>
                {this._gerErrors()}
                {this.state.status}
              </div>
              <p className="ms-font-l ms-fontColor-white">News Items</p>
              <div>
                <div className={styles.row}>
                  <div className={styles.left}>Id</div>
                  <div className={styles.right}>Title</div>
                </div>
                {
                  this.state.items.map((item, idx) => {
                    return(
                      <div className={styles.row}>
                        <div className={styles.left}>{item.Id}</div>
                        <div className={styles.right}>{item.Title}</div>
                      </div>
                    );
                  })
                }
              </div>
            </div>
          </div>
        </div>
    );
  }


  private async _readItems(listName: string): Promise<void> {
    try {
      const web: Web = new Web(this.props.pageContext.web.absoluteUrl);
      const items: INewsItem[] = await web.lists
      .getByTitle(listName)
      .items
      .select("Id","Title")
      .usingCaching()
      .get();

      const status: string = "Loaded news items";
      Log.verbose(LOG_SOURCE, "Items loaded.");
      Log.info(LOG_SOURCE, `List name parameter: ${listName}`);
      Log.verbose(LOG_SOURCE, JSON.stringify(items, undefined, 2));
      this.setState({ ...this.state, items, status});

    } catch (error) {
      this.setState({ ...this.state, errors: [...this.state.errors, error] });
      Log.error(LOG_SOURCE, error);
    }
  }

  private _gerErrors() {
    return this.state.errors.length > 0
      ?
      <div style={{ color: "orangered" }} >
        <div>Errors:</div>
        {
          this.state.errors.map((item, idx) => {
            return (<div key={idx} >{JSON.stringify(item)}</div>);
          })
        }
      </div>
      : null;
  }
}
