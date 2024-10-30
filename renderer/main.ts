import "./style.css";
import { setupCounter } from "./counter.ts";

setupCounter(document.querySelector<HTMLButtonElement>("#counter")!);

// Use contextBridge
window.ipcRenderer.on("main-process-message", (_event, message) => {
	console.log(message);
});
