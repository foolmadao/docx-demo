import { Component, OnInit } from '@angular/core';
import * as docxtemplater from 'docxtemplater';
import * as inspect from 'docxtemplater/js/inspect-module';
import * as PizZip from 'pizzip';
import * as pizzipUtils from 'pizzip/utils';
import * as FileSaver from 'file-saver';
@Component({
  selector: 'app-root',
  templateUrl: './app.component.html'
})
export class AppComponent {
  title = 'doc-test';

  loadFile(url, callback) {
    pizzipUtils.getBinaryContent(url, callback);
  }
  generate() {
      this.loadFile('../assets/tag-example.docx', function(error, content) {

          if (error) { throw error; }
          const zip = PizZip(content);

          const doc = new docxtemplater().loadZip(zip);

          const iModule = inspect();
          doc.attachModule(iModule);
          doc.render(); // doc.compile can also be used to avoid having runtime errors
          const tags = iModule.getAllTags();
          console.log(tags);

          doc.setData({
              first_name: 'John',
              last_name: 'Doe',
              phone: '0652455478',
              description: 'New Website',
              products: [
                {
                  title: 'Duk',
                  name: 'DukSoftware',
                  reference: 'DS0'
                },
                {
                  title: 'Tingerloo',
                  name: 'Tingerlee',
                  reference: 'T00'
                }
              ]
          });
          try {
            doc.render();
          } catch (error) {
              const e = {
                  message: error.message,
                  name: error.name,
                  stack: error.stack,
                  properties: error.properties,
              };
              console.log(JSON.stringify({error: e}));
              // The error thrown here contains additional information when logged with JSON.stringify (it contains a property object).
              throw error;
          }
          const out = doc.getZip().generate({
            type: 'blob'
          });
          FileSaver.saveAs(out, 'output.docx');
      });
  }

}
