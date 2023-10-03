/******/ (function() { // webpackBootstrap
/******/ 	var __webpack_modules__ = ({

/***/ "./src/commands/commands.ts":
/*!**********************************!*\
  !*** ./src/commands/commands.ts ***!
  \**********************************/
/***/ (function(__unused_webpack_module, exports, __webpack_require__) {

/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */
/* global global, Office, self, window */
Office.onReady(function () {
  // If needed, Office.js is ready to be called
});
/**
 * Shows a notification when the add-in command is executed.
 * @param event
 */
function action(event) {
  var message = {
    type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
    message: "Performed action.",
    icon: "Icon.80x80",
    persistent: true
  };
  // Show a notification message
  Office.context.mailbox.item.notificationMessages.replaceAsync("action", message);
  // Be sure to indicate when the add-in command function is complete
  event.completed();
}
function getGlobal() {
  return typeof self !== "undefined" ? self : typeof window !== "undefined" ? window : typeof __webpack_require__.g !== "undefined" ? __webpack_require__.g : undefined;
}
var g = getGlobal();
// The add-in command functions need to be available in global scope
g.action = action;
void function register() {
  /* react-hot-loader/webpack */var reactHotLoader = typeof reactHotLoaderGlobal !== 'undefined' ? reactHotLoaderGlobal.default : undefined;
  if (!reactHotLoader) {
    return;
  } /* eslint-disable camelcase, no-undef */
  var webpackExports = typeof __webpack_exports__ !== 'undefined' ? __webpack_exports__ : exports; /* eslint-enable camelcase, no-undef */
  if (!webpackExports) {
    return;
  }
  if (typeof webpackExports === 'function') {
    reactHotLoader.register(webpackExports, 'module.exports', "U:\\1AMicrosoftAddins\\YoGenerator\\NtTS ReactOfficeAddin\\src\\commands\\commands.ts");
    return;
  } /* eslint-disable no-restricted-syntax */
  for (var key in webpackExports) {
    /* eslint-enable no-restricted-syntax */if (!Object.prototype.hasOwnProperty.call(webpackExports, key)) {
      continue;
    }
    var namedExport = void 0;
    try {
      namedExport = webpackExports[key];
    } catch (err) {
      continue;
    }
    reactHotLoader.register(namedExport, key, "U:\\1AMicrosoftAddins\\YoGenerator\\NtTS ReactOfficeAddin\\src\\commands\\commands.ts");
  }
}();

/***/ })

/******/ 	});
/************************************************************************/
/******/ 	// The require scope
/******/ 	var __webpack_require__ = {};
/******/ 	
/************************************************************************/
/******/ 	/* webpack/runtime/global */
/******/ 	!function() {
/******/ 		__webpack_require__.g = (function() {
/******/ 			if (typeof globalThis === 'object') return globalThis;
/******/ 			try {
/******/ 				return this || new Function('return this')();
/******/ 			} catch (e) {
/******/ 				if (typeof window === 'object') return window;
/******/ 			}
/******/ 		})();
/******/ 	}();
/******/ 	
/************************************************************************/
/******/ 	
/******/ 	// startup
/******/ 	// Load entry module and return exports
/******/ 	// This entry module is referenced by other modules so it can't be inlined
/******/ 	var __webpack_exports__ = {};
/******/ 	__webpack_modules__["./src/commands/commands.ts"](0, __webpack_exports__, __webpack_require__);
/******/ 	
/******/ })()
;
//# sourceMappingURL=commands.js.map