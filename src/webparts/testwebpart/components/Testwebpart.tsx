import * as React from "react";
import styles from "./Testwebpart.module.scss";
import type { ITestwebpartProps } from "./ITestwebpartProps";
import { escape } from "@microsoft/sp-lodash-subset";
import { Web } from "sp-pnp-js";

export default class Testwebpart extends React.Component<
  ITestwebpartProps,
  {}
> {
  public render(): React.ReactElement<ITestwebpartProps> {
    const {
      description,
      isDarkTheme,
      environmentMessage,
      hasTeamsContext,
      userDisplayName,
    } = this.props;

    return (
      <section
        className={`${styles.testwebpart} ${
          hasTeamsContext ? styles.teams : ""
        }`}
      >
        <div className={styles.welcome}>
          <img
            alt=""
            src={
              isDarkTheme
                ? require("../assets/welcome-dark.png")
                : require("../assets/welcome-light.png")
            }
            className={styles.welcomeImage}
          />
          <h2>Well done, {escape(userDisplayName)}!</h2>
          <div>{environmentMessage}</div>
          <div>
            Web part property value: <strong>{escape(description)}</strong>
          </div>
        </div>
        <div>
          <h3>Welcome to SharePoint Framework!</h3>
          <p>
            The SharePoint Framework (SPFx) is a extensibility model for
            Microsoft Viva, Microsoft Teams and SharePoint. It&#39;s the easiest
            way to extend Microsoft 365 with automatic Single Sign On, automatic
            hosting and industry standard tooling.
          </p>
          <h4>Learn more about SPFx development:</h4>
          <ul className={styles.links}>
            <li>
              <a href="https://aka.ms/spfx" target="_blank" rel="noreferrer">
                SharePoint Framework Overview
              </a>
            </li>
            <li>
              <a
                href="https://aka.ms/spfx-yeoman-graph"
                target="_blank"
                rel="noreferrer"
              >
                Use Microsoft Graph in your solution
              </a>
            </li>
            <li>
              <a
                href="https://aka.ms/spfx-yeoman-teams"
                target="_blank"
                rel="noreferrer"
              >
                Build for Microsoft Teams using SharePoint Framework
              </a>
            </li>
            <li>
              <a
                href="https://aka.ms/spfx-yeoman-viva"
                target="_blank"
                rel="noreferrer"
              >
                Build for Microsoft Viva Connections using SharePoint Framework
              </a>
            </li>
            <li>
              <a
                href="https://aka.ms/spfx-yeoman-store"
                target="_blank"
                rel="noreferrer"
              >
                Publish SharePoint Framework applications to the marketplace
              </a>
            </li>
            <li>
              <a
                href="https://aka.ms/spfx-yeoman-api"
                target="_blank"
                rel="noreferrer"
              >
                SharePoint Framework API reference
              </a>
            </li>
            <li>
              <a href="https://aka.ms/m365pnp" target="_blank" rel="noreferrer">
                Microsoft 365 Developer Community
              </a>
            </li>
          </ul>
        </div>
      </section>
    );
  }

  componentDidMount(): void {
    console.log("Hello world!");
    this.getListData();
  }

  private getListData = async () => {
    let w = new Web("https://366pitechnologies.sharepoint.com/sites/FloCard/");
    const list = await w.lists
      .getByTitle("TestList")
      .items.select("Title, CanDelete")
      .filter("CanDelete eq false")
      .getAll();
    console.log(list);
  };
}
