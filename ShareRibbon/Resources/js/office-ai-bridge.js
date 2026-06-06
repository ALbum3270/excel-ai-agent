/**
 * office-ai-bridge.js
 * Centralizes communication from the WebView UI to the VB.NET host.
 */
(function (window) {
    'use strict';

    function normalizeMessage(typeOrMessage, payload) {
        if (typeof typeOrMessage === 'object' && typeOrMessage !== null) {
            return typeOrMessage;
        }

        var message = payload || {};
        message.type = typeOrMessage;
        return message;
    }

    function post(typeOrMessage, payload) {
        var message = normalizeMessage(typeOrMessage, payload);

        try {
            if (window.chrome && window.chrome.webview) {
                window.chrome.webview.postMessage(message);
                return true;
            }

            if (window.vsto) {
                if (typeof window.vsto.sendMessage === 'function') {
                    window.vsto.sendMessage(JSON.stringify(message));
                    return true;
                }

                if (typeof window.vsto.postMessage === 'function') {
                    window.vsto.postMessage(message);
                    return true;
                }
            }
        } catch (error) {
            console.error('[OfficeAiBridge] post failed:', error);
            return false;
        }

        console.error('[OfficeAiBridge] no supported host bridge found');
        return false;
    }

    window.officeAi = window.officeAi || {};
    window.officeAi.post = post;
})(window);
