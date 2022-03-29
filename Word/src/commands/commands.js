/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global global, Office, self, window */
(function () {
  Office.initialize = function (reason) {
      // If you need to initialize something you can do so here.
  };
})();
function writeText(event) {

  // Implement your custom code here. The following code is a simple example.  
  Office.context.document.setSelectedDataAsync("ExecuteFunction works. Button ID=" + event.source.id,
      function (asyncResult) {
          var error = asyncResult.error;
          if (asyncResult.status === Office.AsyncResultStatus.Failed) {
              // Show error message.
          }
          else {
              // Show success message.
          }
      });

  // Calling event.completed is required. event.completed lets the platform know that processing has completed.
  event.completed();
}