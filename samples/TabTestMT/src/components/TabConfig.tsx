import { useContext } from "react";
import { TeamsFxContext } from "./Context";
import { app, pages } from '@microsoft/teams-js';


export default function TabConfig() {
  const { themeString } = useContext(TeamsFxContext);
  app.initialize();
  pages.config.registerOnSaveHandler((saveEvent) => {
    const url = `${window.location.origin}/index.html#/tab`;
    pages.config.setConfig({
      "suggestedDisplayName": "Tab",
      "entityId": "tab",
      "contentUrl": url,
      "websiteUrl": url
    });
    saveEvent.notifySuccess();
  });
  pages.config.setValidityState(true);

  return (
    <div
      className={themeString === "default" ? "light" : themeString === "dark" ? "dark" : "contrast"}
    >
      There's nothing to configure here so please just click save
    </div>
  );
}
