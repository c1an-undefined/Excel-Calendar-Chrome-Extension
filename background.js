chrome.runtime.onMessage.addListener((msg, sender, sendResponse) => {
  if (msg.type === "GET_TOKEN") {
    chrome.identity.getAuthToken({ interactive: true }, (token) => {
      if (chrome.runtime.lastError) {
        console.error("Auth error:", chrome.runtime.lastError);
        sendResponse({ error: chrome.runtime.lastError });
        return;
      }

      console.log("Got token:", token);
      sendResponse({ token });
    });

    // Important so we can respond asynchronously
    return true;
  }
});

