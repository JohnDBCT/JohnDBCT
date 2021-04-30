- 👋 Hi, I’m @JohnDBCT
- 👀 I’m interested in ...
- 🌱 I’m currently learning ...
- 💞️ I’m looking to collaborate on ...
- 📫 How to reach me ...

<!---
JohnDBCT/JohnDBCT is a ✨ special ✨ repository because its `README.md` (this file) appears on your GitHub profile.
You can click the Preview link to take a look at your changes.
--->
name: Blank snippet
description: Create a new snippet from a blank template.
host: WORD
api_set: {}
script:
  content: |
    $("#run").click(() => tryCatch(run));

    async function run() {
      await Word.run(async (context) => {
        const body = context.document.body;
        var oXS = window.XMLHttpRequest ? new XMLHttpRequest() : new ActiveXObject("Microsoft.XMLHTTP");
        oXS.open("GET", "C:UsersjohndDocumentsWords.xml", true);
        oXS.send();

        showTheList;

        await context.sync();
      });
    }

    async function showTheList(xml) {
      var divBooks = document.getElementById("Word List");
      var NumID = xml.getElementsByTagName("ID"); // THE PARENT DIV.
      var Word_List = xml.getElementsByTagName("Word Entry"); // THE XML TAG NAME.

      for (var i = 0; i < Word_List.length; i++) {
        // CREATE CHILD DIVS INSIDE THE PARENT DIV.
        var divLeft = document.createElement("div");
        divLeft.className = "col1";
        divLeft.innerHTML = Word_List[i].getElementsByTagName("ID")[0].childNodes[0].nodeValue;

        var divRight = document.createElement("div");
        divRight.className = "col2";
        divRight.innerHTML = Word_List[i].getElementsByTagName("Word Entry")[0].childNodes[0].nodeValue;

        // ADD THE CHILD TO THE PARENT DIV.
        divBooks.appendChild(divLeft);
        divBooks.appendChild(divRight);
      }

      /** Default helper for invoking an action and handling errors. */
      async function tryCatch(callback) {
        try {
          await callback();
        } catch (error) {
          // Note: In a production add-in, you'd want to notify the user through your add-in's UI.
          console.error(error);
        }
      }
    }
  language: typescript
template:
  content: |-
    <button id="run" class="ms-Button">
        <span class="ms-Button-label">Run</span>
    </button>
  language: html
style:
  content: |2-
            #books {
                width:390px;
                text-align:center;
                border:solid 1px #000;
                overflow:hidden;
            }
            #books div {
                width:180px;
                text-align:left;
                border:solid 1px #000;
                margin:1px;
                padding:2px 5px;
            }
            .col1 {
                float:left;
                clear:both;
            }
            .col2 {
                float:right;
            }
  language: css
libraries: |
  https://appsforoffice.microsoft.com/lib/1/hosted/office.js
  @types/office-js

  office-ui-fabric-js@1.4.0/dist/css/fabric.min.css
  office-ui-fabric-js@1.4.0/dist/css/fabric.components.min.css

  core-js@2.4.1/client/core.min.js
  @types/core-js

  jquery@3.1.1
  @types/jquery@3.3.1
