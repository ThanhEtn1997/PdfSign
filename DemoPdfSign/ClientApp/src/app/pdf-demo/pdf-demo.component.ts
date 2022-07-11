import { Component, OnInit } from '@angular/core';
import SignaturePad from 'signature_pad';
import { HttpClient, HttpHeaders } from '@angular/common/http';

@Component({
  selector: 'app-pdf-demo',
  templateUrl: './pdf-demo.component.html',
  styleUrls: ['./pdf-demo.component.css']
})
export class PdfDemoComponent implements OnInit {

  signaturePad: SignaturePad = null;
  constructor(private http: HttpClient) { }

  ngOnInit() {
    this.createSign();
  }

  createSign() {
    const canvas = document.querySelector("canvas");
    this.signaturePad = new SignaturePad(canvas);
    //signaturePad.on();
  }

  async sign() {
    const data = this.signaturePad.toDataURL(); // save image as PNG
    console.log(data);

    let res = await this.signService({
      image: data
    });
    console.debug(res);
    window.open("http://localhost:32273/" + res);
  }

  download() {
    const data = this.signaturePad.toDataURL(); // save image as PNG
    console.log(data);
    this.debugBase64(data);
  }

  clear() {
    // Clears the canvas
    this.signaturePad.clear();
  }

  debugBase64(base64URL) {
    var win = window.open();
    win.document.write('<iframe src="' + base64URL + '" frameborder="0" style="border:0; top:0px; left:0px; bottom:0px; right:0px; width:100%; height:100%;" allowfullscreen></iframe>');
  }

  async signService(data: any): Promise<any> {
    let promise = new Promise((resolve, reject) => {
      this.http.post('http://localhost:32273/api/PdfSign/sign2', data, {responseType: "text"}).subscribe(res => {
        resolve(res);
      }, err => {
        reject(err);
      });
    });

    return promise;
  }
}
