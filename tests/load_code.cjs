'use strict';
const fs = require('fs');
const path = require('path');

// Nạp ../Code.gs như một module Node (gọi SAU buildMockEnv() để các service GAS giả lập đã có sẵn).
module.exports = function loadCode() {
  const code = fs.readFileSync(path.join(__dirname, '..', 'Code.gs'), 'utf8');
  const names = [...code.matchAll(/^function (\w+)/gm)].map((m) => m[1])
    .concat(['SS_IDS', 'MAU_WORD_FOLDER_ID', 'OWN_HEADERS']);
  const mod = new Function(code + '\nreturn {' + names.join(',') + '};')();

  const mauWordFolder = global.DriveApp.getFolderById('ROOT_MAU_WORD');
  const orig = global.DriveApp.getFolderById;
  global.DriveApp.getFolderById = (id) => (id === mod.MAU_WORD_FOLDER_ID ? mauWordFolder : orig(id));
  return mod;
};
