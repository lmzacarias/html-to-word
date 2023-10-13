const ejs = require('ejs');
const fs = require('fs');
const readFile = require('fs/promises').readFile;
const htmlDocx = require('html-docx-js');

const ejsTemplate = `
<!DOCTYPE html>
<html lang="en">
  <head>
    <meta charset="utf-8">
    <title>N&oacute;mina de estudiantes</title>
    <style type="text/css">
      <%= css %>
    </style>
  </head>
  <body>
    <table class="pdf_main_table">
      <thead class="bordered">
        <tr>
              <th class="width-4">No.</th>
              <th class="text-left full-name width-20">&nbsp;ESTUDIANTES</th>
              <% for (column = 0; column < columns; column ++) { %>
                <th></th>
              <% } %>
            </tr>
        </thead>
      </table>

    
    <table class="pdf_main_table">
      <tbody class="bordered">
        <% students.forEach(function(student, index) { %>
          <tr>
            <td class="text-center width-4"><%= index+1 %></td>
            <td class="text-left upper-case full-name width-20">&nbsp;<%=student.relational_data.name.show%></td>
            <% for (column = 0; column < columns; column ++) { %>
              <td></td>
            <% } %>
          </tr>
        <% }) %>
      </tbody>
    </table>
  </body>
</html>
`;

async function init() {
  const data = {
    title: 'Mi Título',
    content: 'Este es el contenido de mi documento.',
  };

  const data2 = [
    {
      relational_data: {
        name: {
          show: "Allerdyce Pyro, John",
          order: "allerdyce pyro, john"
        }
      }
    },
    {
      relational_data: {
        name: {
          show: "Allerdyce Pyro, John",
          order: "allerdyce pyro, john"
        }
      }
    },
    {
      relational_data: {
        name: {
          show: "Allerdyce Pyro, John",
          order: "allerdyce pyro, john"
        }
      }
    },
    {
      relational_data: {
        name: {
          show: "Allerdyce Pyro, John",
          order: "allerdyce pyro, john"
        }
      }
    },
    {
      relational_data: {
        name: {
          show: "Allerdyce Pyro, John",
          order: "allerdyce pyro, john"
        }
      }
    },
    {
      relational_data: {
        name: {
          show: "Allerdyce Pyro, John",
          order: "allerdyce pyro, john"
        }
      }
    },
    {
      relational_data: {
        name: {
          show: "Allerdyce Pyro, John",
          order: "allerdyce pyro, john"
        }
      }
    },
    {
      relational_data: {
        name: {
          show: "Allerdyce Pyro, John",
          order: "allerdyce pyro, john"
        }
      }
    },
    {
      relational_data: {
        name: {
          show: "Allerdyce Pyro, John",
          order: "allerdyce pyro, john"
        }
      }
    },

  ]

  const css = await readFile(__dirname + '/styles.css');
  const outputPath = 'output.docx';

  convertEjsToWord(ejsTemplate, { css: css.toString(), students: data2, columns: 1 }, outputPath).then(() => {
    console.log('Archivo .docx generado correctamente.');
  });
}

async function convertEjsToWord(template, data, outputPath) {
  const renderedHtml = ejs.render(template, data);
  const buffer = htmlDocx.asBlob(renderedHtml);
  fs.writeFileSync(outputPath, buffer);
}

init()
