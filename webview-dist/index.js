function t(t,e,n,r){return new(n||(n=Promise))((function(o,i){function s(t){try{a(r.next(t))}catch(t){i(t)}}function u(t){try{a(r.throw(t))}catch(t){i(t)}}function a(t){var e;t.done?o(t.value):(e=t.value,e instanceof n?e:new n((function(t){t(e)}))).then(s,u)}a((r=r.apply(t,e||[])).next())}))}function e(t,e){var n,r,o,i,s={label:0,sent:function(){if(1&o[0])throw o[1];return o[1]},trys:[],ops:[]};return i={next:u(0),throw:u(1),return:u(2)},"function"==typeof Symbol&&(i[Symbol.iterator]=function(){return this}),i;function u(i){return function(u){return function(i){if(n)throw new TypeError("Generator is already executing.");for(;s;)try{if(n=1,r&&(o=2&i[0]?r.return:i[0]?r.throw||((o=r.return)&&o.call(r),0):r.next)&&!(o=o.call(r,i[1])).done)return o;switch(r=0,o&&(i=[2&i[0],o.value]),i[0]){case 0:case 1:o=i;break;case 4:return s.label++,{value:i[1],done:!1};case 5:s.label++,r=i[1],i=[0];continue;case 7:i=s.ops.pop(),s.trys.pop();continue;default:if(!(o=s.trys,(o=o.length>0&&o[o.length-1])||6!==i[0]&&2!==i[0])){s=0;continue}if(3===i[0]&&(!o||i[1]>o[0]&&i[1]<o[3])){s.label=i[1];break}if(6===i[0]&&s.label<o[1]){s.label=o[1],o=i;break}if(o&&s.label<o[2]){s.label=o[2],s.ops.push(i);break}o[2]&&s.ops.pop(),s.trys.pop();continue}i=e.call(t,s)}catch(t){i=[6,t],r=0}finally{n=o=0}if(5&i[0])throw i[1];return{value:i[0]?i[1]:void 0,done:!0}}([i,u])}}}var n=function(){return(n=Object.assign||function(t){for(var e,n=1,r=arguments.length;n<r;n++)for(var o in e=arguments[n])Object.prototype.hasOwnProperty.call(e,o)&&(t[o]=e[o]);return t}).apply(this,arguments)};function r(t,e){void 0===e&&(e=!1);var n=window.crypto.getRandomValues(new Uint32Array(1))[0],r="_".concat(n);return Object.defineProperty(window,r,{value:function(n){return e&&Reflect.deleteProperty(window,r),null==t?void 0:t(n)},writable:!1,configurable:!0}),n}function o(t,e){return void 0===e&&(e={}),function(t,e,n,r){return new(n||(n=Promise))((function(o,i){function s(t){try{a(r.next(t))}catch(t){i(t)}}function u(t){try{a(r.throw(t))}catch(t){i(t)}}function a(t){var e;t.done?o(t.value):(e=t.value,e instanceof n?e:new n((function(t){t(e)}))).then(s,u)}a((r=r.apply(t,e||[])).next())}))}(this,void 0,void 0,(function(){return function(t,e){var n,r,o,i,s={label:0,sent:function(){if(1&o[0])throw o[1];return o[1]},trys:[],ops:[]};return i={next:u(0),throw:u(1),return:u(2)},"function"==typeof Symbol&&(i[Symbol.iterator]=function(){return this}),i;function u(i){return function(u){return function(i){if(n)throw new TypeError("Generator is already executing.");for(;s;)try{if(n=1,r&&(o=2&i[0]?r.return:i[0]?r.throw||((o=r.return)&&o.call(r),0):r.next)&&!(o=o.call(r,i[1])).done)return o;switch(r=0,o&&(i=[2&i[0],o.value]),i[0]){case 0:case 1:o=i;break;case 4:return s.label++,{value:i[1],done:!1};case 5:s.label++,r=i[1],i=[0];continue;case 7:i=s.ops.pop(),s.trys.pop();continue;default:if(!((o=(o=s.trys).length>0&&o[o.length-1])||6!==i[0]&&2!==i[0])){s=0;continue}if(3===i[0]&&(!o||i[1]>o[0]&&i[1]<o[3])){s.label=i[1];break}if(6===i[0]&&s.label<o[1]){s.label=o[1],o=i;break}if(o&&s.label<o[2]){s.label=o[2],s.ops.push(i);break}o[2]&&s.ops.pop(),s.trys.pop();continue}i=e.call(t,s)}catch(t){i=[6,t],r=0}finally{n=o=0}if(5&i[0])throw i[1];return{value:i[0]?i[1]:void 0,done:!0}}([i,u])}}}(this,(function(o){return[2,new Promise((function(o,i){var s=r((function(t){o(t),Reflect.deleteProperty(window,"_".concat(u))}),!0),u=r((function(t){i(t),Reflect.deleteProperty(window,"_".concat(s))}),!0);window.__TAURI_IPC__(n({cmd:t,callback:s,error:u},e))}))]}))}))}Object.freeze({__proto__:null,transformCallback:r,invoke:o,convertFileSrc:function(t,e){void 0===e&&(e="asset");var n=encodeURIComponent(t);return navigator.userAgent.includes("Windows")?"https://".concat(e,".localhost/").concat(n):"".concat(e,"://").concat(n)}});var i=function(){function n(t,e){this.path=t,this.sheetName=e}return n.prototype.close=function(){return t(this,void 0,void 0,(function(){return e(this,(function(t){switch(t.label){case 0:return[4,o("plugin:spreadsheet|close_xlsx",{path:this.path})];case 1:return[2,t.sent()]}}))}))},n.closeAll=function(){return t(this,void 0,void 0,(function(){return e(this,(function(t){switch(t.label){case 0:return[4,o("plugin:spreadsheet|close_all_xlsx")];case 1:return[2,t.sent()]}}))}))},n.prototype.copySheet=function(n,r){return t(this,void 0,void 0,(function(){return e(this,(function(t){switch(t.label){case 0:return[4,o("plugin:spreadsheet|copy_sheet",{path:this.path,sourceSheetName:r||this.sheetName,targetSheetName:n})];case 1:return[2,t.sent()]}}))}))},n.prototype.create=function(){return t(this,void 0,void 0,(function(){return e(this,(function(t){switch(t.label){case 0:return[4,o("plugin:spreadsheet|new_xlsx",{path:this.path})];case 1:return[2,t.sent()]}}))}))},n.prototype.getSheetColumn=function(){return t(this,void 0,void 0,(function(){return e(this,(function(t){switch(t.label){case 0:return[4,o("plugin:spreadsheet|get_sheet_highest_column",{path:this.path,sheetName:this.sheetName})];case 1:return[2,t.sent()]}}))}))},n.prototype.getSheetRange=function(){return t(this,void 0,void 0,(function(){return e(this,(function(t){switch(t.label){case 0:return[4,o("plugin:spreadsheet|get_sheet_highest_column_and_row",{path:this.path,sheetName:this.sheetName})];case 1:return[2,t.sent()]}}))}))},n.prototype.getSheetRow=function(){return t(this,void 0,void 0,(function(){return e(this,(function(t){switch(t.label){case 0:return[4,o("plugin:spreadsheet|get_sheet_highest_row",{path:this.path,sheetName:this.sheetName})];case 1:return t.sent(),[2]}}))}))},n.prototype.getValue=function(n){return t(this,void 0,void 0,(function(){return e(this,(function(t){switch(t.label){case 0:return[4,o("plugin:spreadsheet|get_value_by_column_and_row",{path:this.path,sheetName:this.sheetName,local:n})];case 1:return[2,t.sent()]}}))}))},n.list=function(){return t(this,void 0,void 0,(function(){return e(this,(function(t){switch(t.label){case 0:return[4,o("plugin:spreadsheet|list_xlsx")];case 1:return[2,t.sent()]}}))}))},n.prototype.newSheet=function(n){return t(this,void 0,void 0,(function(){return e(this,(function(t){switch(t.label){case 0:return[4,o("plugin:spreadsheet|new_sheet",{path:this.path,sheetName:n})];case 1:return[2,t.sent()]}}))}))},n.prototype.read=function(){return t(this,void 0,void 0,(function(){return e(this,(function(t){switch(t.label){case 0:return[4,o("plugin:spreadsheet|read_xlsx",{path:this.path})];case 1:return[2,t.sent()]}}))}))},n.prototype.removeRow=function(n,r){return t(this,void 0,void 0,(function(){return e(this,(function(t){switch(t.label){case 0:return[4,o("plugin:spreadsheet|remove_row",{start:n,length:r,path:this.path,sheetName:this.sheetName})];case 1:return[2,t.sent()]}}))}))},n.prototype.setValue=function(n,r){return t(this,void 0,void 0,(function(){return e(this,(function(t){switch(t.label){case 0:return[4,o("plugin:spreadsheet|set_value_by_column_and_row",{path:this.path,sheetName:this.sheetName,local:n,value:r})];case 1:return[2,t.sent()]}}))}))},n.prototype.write=function(){return t(this,void 0,void 0,(function(){return e(this,(function(t){switch(t.label){case 0:return[4,o("plugin:spreadsheet|write_xlsx",{path:this.path})];case 1:return[2,t.sent()]}}))}))},n}();export{i as Spreadsheet};
