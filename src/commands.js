Office.onReady(() => {});

const ARCHIVE_FOLDER_NAMES = ["Archiv", "Archive"];

function archiveAndReply(event) {
  const item = Office.context.mailbox.item;
  const payload = {
    archiveItemId: item.itemId,
    archiveConversationId: item.conversationId || ""
  };
  Office.addin
    .setStartupVariables(payload)
    .then(function () {
      item.displayReplyForm({});
      event.completed();
    })
    .catch(function (error) {
      console.error("setStartupVariables failed:", error);
      item.displayReplyForm({});
      event.completed();
    });
}

function onMessageSendHandler(event) {
  Office.addin
    .getStartupVariables()
    .then(function (variables) {
      const archiveItemId = variables && variables.archiveItemId;
      const archiveConversationId = variables && variables.archiveConversationId;

      if (!archiveItemId) {
        event.completed({ allowEvent: true, cancelLabel: "Weiter" });
        return;
      }

      const currentItem = Office.context.mailbox.item;
      const sameConversation =
        !archiveConversationId ||
        currentItem.conversationId === archiveConversationId;

      if (!sameConversation) {
        event.completed({ allowEvent: true, cancelLabel: "Weiter" });
        return;
      }

      Office.addin.setStartupVariables({ archiveItemId: "" }).then(function () {
        moveToArchiveFolder(archiveItemId, function () {
          event.completed({ allowEvent: true, cancelLabel: "Weiter" });
        });
      });
    })
    .catch(function (error) {
      console.error("getStartupVariables failed:", error);
      event.completed({ allowEvent: true, cancelLabel: "Weiter" });
    });
}

function moveToArchiveFolder(itemId, done) {
  resolveArchiveFolderId(
    function (folderId) {
      moveItem(itemId, folderId, done, done);
    },
    function () {
      console.error("Archiv-Ordner nicht gefunden. Element wird nicht verschoben.");
      done();
    }
  );
}

function resolveArchiveFolderId(success, failure) {
  Office.context.mailbox.makeEwsRequestAsync(
    getFindFolderRequest(ARCHIVE_FOLDER_NAMES),
    function (result) {
      if (result.status !== "succeeded") return failure();
      try {
        const doc = new DOMParser().parseFromString(result.value, "text/xml");
        const folderId = doc.getElementsByTagName("t:FolderId")[0];
        if (!folderId) return failure();
        success(folderId.getAttribute("Id"));
      } catch (e) {
        failure();
      }
    }
  );
}

function getFindFolderRequest(names) {
  const folderNames = names
    .map(function (n) {
      return "<t:Contains ContainmentMode=" + '"ExactPhrase"' + " ContainmentComparison=" + '"IgnoreCase"' + ">"
        + "<t:FieldURI FieldURI=" + '"folder:DisplayName"' + "/>"
        + "<t:Constant Value=\"" + escapeXml(n) + "\"/>"
        + "</t:Contains>";
    })
    .join("");

  return (
    '<?xml version="1.0" encoding="utf-8"?>' +
    '<soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" ' +
    'xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages" ' +
    'xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types" ' +
    'xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">' +
    "<soap:Header><t:RequestServerVersion Version=\"Exchange2013\"/></soap:Header>" +
    "<soap:Body>" +
    '<m:FindFolder Traversal="Deep">' +
    "<m:FolderShape><t:BaseShape>Default</t:BaseShape></m:FolderShape>" +
    "<m:Restriction>" +
    '<t:Or>' + folderNames + "</t:Or>" +
    "</m:Restriction>" +
    "<m:ParentFolderIds><t:DistinguishedFolderId Id=\"msgfolderroot\"/></m:ParentFolderIds>" +
    "</m:FindFolder>" +
    "</soap:Body></soap:Envelope>"
  );
}

function moveItem(itemId, folderId, success, failure) {
  Office.context.mailbox.makeEwsRequestAsync(
    getMoveItemRequest(itemId, folderId),
    function (result) {
      if (result.status === "succeeded") return success();
      console.error("MoveItem fehlgeschlagen:", result.error);
      failure();
    }
  );
}

function getMoveItemRequest(itemId, folderId) {
  return (
    '<?xml version="1.0" encoding="utf-8"?>' +
    '<soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" ' +
    'xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages" ' +
    'xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types" ' +
    'xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">' +
    "<soap:Header><t:RequestServerVersion Version=\"Exchange2013\"/></soap:Header>" +
    "<soap:Body>" +
    "<m:MoveItem>" +
    "<m:ToFolderId><t:FolderId Id=\"" + escapeXml(folderId) + "\"/></m:ToFolderId>" +
    "<m:ItemIds><t:ItemId Id=\"" + escapeXml(itemId) + "\"/></m:ItemIds>" +
    "</m:MoveItem>" +
    "</soap:Body></soap:Envelope>"
  );
}

function escapeXml(value) {
  return String(value)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&apos;");
}

Office.actions.associate("archiveAndReply", archiveAndReply);
Office.actions.associate("onMessageSendHandler", onMessageSendHandler);
