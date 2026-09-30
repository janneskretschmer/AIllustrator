/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office, PowerPoint, fetch, localStorage */

const STORAGE_KEY = "aillustrator-openai-key";

Office.onReady((info) => {
  if (info.host === Office.HostType.PowerPoint) {
    document.getElementById("sideload-msg")!.style.display = "none";
    document.getElementById("app-body")!.style.display = "flex";
    
    // Bind buttons
    document.getElementById("run")!.onclick = runPowerPoint;
    document.getElementById("save-key-btn")!.onclick = handleSaveApiKey;

    // Load existing API Key from LocalStorage on load
    loadSavedApiKey();
  }
});

interface ShapeBounds {
  left: number;
  top: number;
  width: number;
  height: number;
}

interface ImageConfig {
  size: "1536x1024" | "1024x1024" | "1024x1536";
  imgRatio: number;
}

// Safely displays status/error messages in the taskpane
function showStatusMessage(message: string, isError: boolean = true) {
  const msgArea = document.getElementById("message-area");
  if (msgArea) {
    msgArea.textContent = message;
    msgArea.style.color = isError ? "#d13438" : "#107c10";
  }
}

// Loads the stored key into the input box on launch
function loadSavedApiKey() {
  const savedKey = localStorage.getItem(STORAGE_KEY);
  if (savedKey) {
    const apiKeyInput = document.getElementById("api-key") as HTMLInputElement;
    if (apiKeyInput) {
      apiKeyInput.value = savedKey;
    }
  }
}

// Handler for the Save button
function handleSaveApiKey() {
  const apiKeyInput = document.getElementById("api-key") as HTMLInputElement;
  const apiKey = apiKeyInput ? apiKeyInput.value.trim() : "";

  if (apiKey) {
    localStorage.setItem(STORAGE_KEY, apiKey);
    showStatusMessage("API Key saved to local storage!", false);
  } else {
    localStorage.removeItem(STORAGE_KEY);
    showStatusMessage("API Key removed from local storage.", false);
  }
}

export async function runPowerPoint() {
  showStatusMessage("", false);

  const apiKeyInput = document.getElementById("api-key") as HTMLInputElement;
  let apiKey = apiKeyInput ? apiKeyInput.value.trim() : "";

  // Fallback to localStorage if input field wasn't updated
  if (!apiKey) {
    apiKey = localStorage.getItem(STORAGE_KEY) || "";
  }

  if (!apiKey) {
    showStatusMessage("Please enter and save your OpenAI API key.");
    return;
  }

  showStatusMessage("Reading slide content...", false);

  let promptText = "";
  let shapePos: ShapeBounds = { left: 0, top: 0, width: 0, height: 0 };

  // 1. Extract text and dimensions from selected text box
  try {
    await PowerPoint.run(async (context: PowerPoint.RequestContext) => {
      const selectedShapes: PowerPoint.ShapeCollection = context.presentation.getSelectedShapes();
      selectedShapes.load("items");
      await context.sync();

      if (selectedShapes.items.length === 0) {
        throw new Error("NoShapeSelected");
      }

      const shape: PowerPoint.Shape = selectedShapes.items[0];
      shape.load("top, left, width, height, textFrame/textRange/text");
      await context.sync();

      shapePos = {
        left: shape.left,
        top: shape.top,
        width: shape.width,
        height: shape.height
      };

      try {
        promptText = shape.textFrame.textRange.text.trim();
        console.log("AIllustrator - Extracted Prompt:", promptText);
      } catch (e) {
        console.warn("AIllustrator: Selected shape does not contain text range.");
      }
    });
  } catch (error: any) {
    if (error.message === "NoShapeSelected") {
      showStatusMessage("Please select a text box on the slide first.");
    } else {
      console.error("AIllustrator - PowerPoint context error:", error);
      showStatusMessage("Failed to read PowerPoint selection.");
    }
    return;
  }

  if (!promptText) {
    showStatusMessage("The selected text box is empty.");
    return;
  }

  // 2. Select image size and calculate aspect-ratio-fitted bounds
  const boxRatio = shapePos.width / shapePos.height;
  const imageConfig = selectImageConfig(boxRatio);
  const fittedBounds = calculateFittedBounds(shapePos, imageConfig.imgRatio);

  // 3. Request image from OpenAI
  showStatusMessage("Generating image from OpenAI... (this takes a few seconds)", false);

  let base64Image = "";
  try {
    base64Image = await generateOpenAIImage(promptText, imageConfig.size, apiKey);
  } catch (e: any) {
    console.error("AIllustrator - OpenAI API Error:", e);
    showStatusMessage(`Generation failed: ${e.message}`);
    return;
  }

  // 4. Overlay the image centered over the text box
  showStatusMessage("Inserting image...", false);

  Office.context.document.setSelectedDataAsync(
    base64Image,
    {
      coercionType: Office.CoercionType.Image,
      imageLeft: fittedBounds.left,
      imageTop: fittedBounds.top,
      imageWidth: fittedBounds.width,
      imageHeight: fittedBounds.height
    },
    (asyncResult: Office.AsyncResult<void>) => {
      if (asyncResult.status === Office.AsyncResultStatus.Failed) {
        console.error("AIllustrator - Failed to insert image: ", asyncResult.error.message);
        showStatusMessage(`Insert failed: ${asyncResult.error.message}`);
      } else {
        showStatusMessage("Image overlay added successfully!", false);
      }
    }
  );
}

function selectImageConfig(boxRatio: number): ImageConfig {
  if (boxRatio > 1.25) {
    return { size: "1536x1024", imgRatio: 1.5 };
  } else if (boxRatio < 0.8) {
    return { size: "1024x1536", imgRatio: 1024 / 1536 };
  } else {
    return { size: "1024x1024", imgRatio: 1.0 };
  }
}

function calculateFittedBounds(box: ShapeBounds, imgRatio: number): ShapeBounds {
  const boxRatio = box.width / box.height;
  let finalWidth: number;
  let finalHeight: number;
  let finalLeft: number;
  let finalTop: number;

  if (boxRatio > imgRatio) {
    finalHeight = box.height;
    finalWidth = box.height * imgRatio;
    finalTop = box.top;
    finalLeft = box.left + (box.width - finalWidth) / 2;
  } else {
    finalWidth = box.width;
    finalHeight = box.width / imgRatio;
    finalLeft = box.left;
    finalTop = box.top + (box.height - finalHeight) / 2;
  }

  return {
    left: finalLeft,
    top: finalTop,
    width: finalWidth,
    height: finalHeight
  };
}

async function generateOpenAIImage(
  prompt: string,
  size: "1536x1024" | "1024x1024" | "1024x1536",
  apiKey: string
): Promise<string> {
  const response = await fetch("https://api.openai.com/v1/images/generations", {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
      Authorization: `Bearer ${apiKey}`
    },
    body: JSON.stringify({
      model: "gpt-image-2.5-flare",
      prompt: prompt,
      n: 1,
      size: size,
      quality: "medium"
    })
  });

  if (!response.ok) {
    const errorData = await response.json().catch(() => ({}));
    const message = errorData?.error?.message || response.statusText;
    throw new Error(`API Error (${response.status}): ${message}`);
  }

  const data = await response.json();
  if (!data?.data?.[0]?.b64_json) {
    throw new Error("No image data returned from OpenAI API.");
  }

  return data.data[0].b64_json;
}