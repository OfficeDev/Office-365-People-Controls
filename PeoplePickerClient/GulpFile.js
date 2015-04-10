﻿var gulp = require('gulp'),
    minifycss = require('gulp-minify-css'),
    concat = require('gulp-concat'),
    uglify = require('gulp-uglify'),
    rename = require('gulp-rename');

gulp.task('minifycss', function () {
    return gulp.src('Office.Controls.PeoplePicker.css')
        .pipe(rename({ suffix: '.min' }))
        .pipe(minifycss())
        .pipe(gulp.dest(''));
});

gulp.task('minifyjs', function () {
    return gulp.src('Office.Controls.PeoplePicker.js')
       // .pipe(concat('main.js'))    //merge all js to main.js
       // .pipe(gulp.dest('minified/js'))    //put main.js to this folder
        .pipe(rename({ suffix: '.min' }))
        .pipe(uglify())
        .pipe(gulp.dest(''));
});

gulp.task('default', function () {
    gulp.start('minifycss', 'minifyjs');
});