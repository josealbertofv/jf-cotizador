# jf-cotizador
It is a software for making quotes for civil construction works and electromechanical construction works


I've made this with javascript language in Google Apps scripts

Instructions to install clasp, in order to work with Google Apps Scripts in vscode IDE: 

https://www.youtube.com/watch?v=4Qlt3p6N0es&t=262s

(I) Instructions to install clasp

(1) You need to make sure to have node js installed:

$ node -v

(2) You need to make sure to have Google Apps API turned on: script.google.com/home/usersettings

(3) Then install clasp:

$ sudo npm install -g @google/clasp

(4) Check clasp version:

$ clasp -v


(II) Instructions to pull the project

(1) Copy de Script ID:

https://script.google.com --> File --> Prject properties --> Info

(2) Log in with google:

$ clasp login

(3) Initialize a node js project inside the folder containing your project. 

$ /path/to/your/folder/ npm init

It will create /path/to/your/folder/package.json

(4) Clone your script inside the folder containing your scripts, for example .../src/:

$ /path/to/your/folder/ clasp clone "Script ID" --rootDir src

It will create
/path/to/your/folder/.clasp.json
/path/to/your/folder/src/appsscript.json


Once you make changes locally, you need to refresh those changes in the actual seerver.

(III) Instructions to apply changes from local to the actual server

$ clasp push

$ clasp push -w    // apply changes automatically while running this app

To exit the app just press Ctrl + C


(IV) Instructions to apply changes from remote

$ clasp pull


(V) To enable the autocomplete tool

Install TypeScript definitions for Apps Script in your project's folder, as it says here https://github.com/google/clasp/blob/master/docs/typescript.md

$ npm i -S @types/google-apps-script
