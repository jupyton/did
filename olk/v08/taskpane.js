
Office.onReady((info) => {
  console.info(`[commands.js::Office.onReady([${JSON.stringify(info)}])]`);


  if (info.host === Office.HostType.Outlook) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    getCategories();
    document.getElementById("apply-categories").onclick = applyCategories;
    document.getElementById("categories-container").onclick = function () {
      clearElement("notification");
    };
    getAppliedCategories();
  }
});


function getCategories() {
  Office.context.mailbox.masterCategories.getAsync((asyncResult) => {
    if (asyncResult.status === Office.AsyncResultStatus.Failed) {
      console.log(asyncResult.error.message);
      return;
    }

    clearElement("categories-container"); 
    let selection = document.createElement("select");
    selection.name = "applicable-categories";
    selection.id = "applicable-categories";
    selection.multiple = true;
    let label = document.createElement("label");
    label.innerHTML =
      "<br/>Select the applicable categories.<br/><br/>Select and hold <kbd>Ctrl</kbd> to choose multiple categories.<br/>";
    label.htmlFor = "applicable-categories";

    asyncResult.value.forEach((category, index) => {
      let displayName = category.displayName;
      if (displayName.includes("Office Add-ins Sample: ")) {
        let option = document.createElement("option");
        option.value = index;
        option.text = category.displayName;
        selection.appendChild(option);
        selection.size++;
      }
    });

    const container = document.getElementById("categories-container");
    container.appendChild(label);
    container.appendChild(selection);
  });
}


function applyCategories() {
  let selectedCategories = getSelectedCategories("applicable-categories");
  if (selectedCategories.length > 0) {
    Office.context.mailbox.item.categories.addAsync(
      selectedCategories,
      (asyncResult) => {
        if (asyncResult.status === Office.AsyncResultStatus.Failed) {
          console.log(asyncResult.error.message);
          return;
        }

        document.getElementById("notification").innerHTML =
          "<i>Selected categories have been applied.</i>";
        getAppliedCategories();
        clearSelection("applicable-categories");
      }
    );
  } else {
    document.getElementById("notification").innerHTML =
      "<i>Select categories to be applied.</i>";
  }
}


function getSelectedCategories(id) {
  let selectedCategories = [];
  for (let category of document.getElementById(id).options) {
    if (category.selected) {
      selectedCategories.push(category.text);
    }
  }

  return selectedCategories;
}


function getAppliedCategories() {
  clearElement("applied-categories-container");

  Office.context.mailbox.item.categories.getAsync((asyncResult) => {
    if (asyncResult.status === Office.AsyncResultStatus.Failed) {
      console.log(asyncResult.error.message);
    }
    
    let categories = asyncResult.value;
    if (categories) {
      let categoryList = document.createElement("ul");
      categories.forEach((category) => {
        let appliedCategory = document.createElement("li");
        appliedCategory.innerText = category.displayName;
        categoryList.appendChild(appliedCategory);
      });

      document
        .getElementById("applied-categories-container")
        .appendChild(categoryList);
    } else {
      let notification = document.createElement("p");
      notification.innerText = "No categories have been applied.";
      document
        .getElementById("applied-categories-container")
        .appendChild(notification);
    }
  });
}


function clearElement(id) {
  document.getElementById(id).innerHTML = "";
}


function clearSelection(id) {
  document.getElementById(id).options.forEach((option) => {
    if (option.selected) {
      option.selected = false;
    }
  });
}