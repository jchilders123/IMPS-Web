(function(){"use strict";const i="https://cdn.jsdelivr.net/pyodide/v0.27.7/full/";let o=null;async function s(t){const e=await fetch(t);if(!e.ok)throw new Error(`Could not load ${t} (${e.status}).`);return new Uint8Array(await e.arrayBuffer())}async function a(t){const e=await fetch(t);if(!e.ok)throw new Error(`Could not load ${t} (${e.status}).`);return e.text()}async function c(t){return o||(o=(async()=>{postMessage({type:"progress",message:"Loading the browser document engine…"});const{loadPyodide:e}=await import(`${i}pyodide.mjs`),n=await e({indexURL:i});await n.loadPackage(["micropip","lxml"]),await n.runPythonAsync(`
import micropip
await micropip.install("python-docx==1.2.0")
`);const[r,p,d]=await Promise.all([a(t.generator),s(t.template),s(t.reference)]);return n.FS.writeFile("/generate_ics206.py",r,{encoding:"utf8"}),n.FS.writeFile("/template.docx",p),n.FS.writeFile("/reference.docx",d),await n.runPythonAsync(`
import importlib.util
spec = importlib.util.spec_from_file_location("imps_generator", "/generate_ics206.py")
imps_generator = importlib.util.module_from_spec(spec)
spec.loader.exec_module(imps_generator)
`),postMessage({type:"progress",message:"Document engine ready."}),n})(),o)}self.addEventListener("message",async t=>{if(t.data?.type==="generate")try{const e=await c(t.data.resources);postMessage({type:"progress",message:"Building the Word ICS-206 on this device…"}),e.globals.set("imps_incident_json",JSON.stringify(t.data.incident)),await e.runPythonAsync(`
import json
incident = json.loads(imps_incident_json)
imps_generator.generate_ics206(
    incident,
    "/template.docx",
    "/output.docx",
    "/reference.docx",
)
`);const n=e.FS.readFile("/output.docx"),r=new Uint8Array(n);postMessage({type:"result",bytes:r.buffer},[r.buffer])}catch(e){postMessage({type:"error",message:e?.message||String(e||"The Word document could not be generated.")})}})})();
