'use strict';

const build = require('@microsoft/sp-build-web');
const { addFastServe } = require("spfx-fast-serve-helpers");

build.addSuppression(`Warning - [sass] The local CSS class 'ms-Grid' is not camelCase and will not be type-safe.`);

const getTasks = build.rig.getTasks;
build.rig.getTasks = function () {
  var result = getTasks.call(build.rig);

  result.set('serve', result.get('serve-deprecated'));

  return result;
};

addFastServe(build);

const gulp = require('gulp');
build.initialize(gulp);
