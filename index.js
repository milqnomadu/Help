import HeadlessPhotopea from "headlessphotopea";
import { readFileSync, writeFileSync, readdirSync } from "fs";
import { dirname } from "path";
import { fileURLToPath } from "url";

const __filename = fileURLToPath(import.meta.url);
const __dirname = dirname(__filename);

let pea = new HeadlessPhotopea({ showBrowser: true, });

let mockup = readFileSync(__dirname + "/mockup.psd");
await pea.addBinaryAsset(mockup); // make sure to use await because it is asynchronous. This is why using a module type package would be better than commonjs here

let imgs = [];
// load all of the images in the images/ folder and populate them into imgs
let imageNames = readdirSync(__dirname + "/images");
for (let n of imageNames) {
    let b64 = readFileSync(__dirname + "/images/" + n, "base64");
    let b64uri = "data:image/png;base64," + b64;
    imgs.push(b64uri);
}

// you lowkey need some photoshop scripting knowledge for this part so make sure to look it up if you haven't already

// call the main document "main"
await pea.runScript(`app.activeDocument.name = "main";`);
// open up the smart object into a new document
await pea.runScript(`app.activeDocument.activeLayer = app.activeDocument.layers.getByName("template");`);
await pea.runScript(`executeAction(stringIDToTypeID("placedLayerEditContents"));`);
// rename the smart object doucment
await pea.runScript(`app.activeDocument.name = "smartObj";`);
// now you can navigate back and forth between main and smartObj :)

let i = 0; // make sure each outputted file has a unique filename

for (let img of imgs) {
    // go to smart object (make sure you add .psd when you use getByName)
    await pea.runScript(`app.activeDocument = app.documents.getByName("smartObj.psd");`);
    // open your image
    await pea.openFromURL(img);
    // get rid of other layers in the smart object (easiest way is to flatten image but this does not always work, you might have to come up with another script for this)
    await pea.runScript(`app.activeDocument.flatten();`);

    // save smart object and go back to main document
    await pea.runScript(`
        app.activeDocument.save();
        app.activeDocument = app.documents.getByName("main.psd");
    `);

    // get output image buffer
    let buff = await pea.exportImage("webp");
    
    // now you can output result as an image
    writeFileSync("output/" + i + ".webp", buff);

    i++; // increment
}

