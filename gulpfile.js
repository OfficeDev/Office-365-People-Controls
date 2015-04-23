var gulp = require('gulp'),
    minifycss = require('gulp-minify-css'),
    concat = require('gulp-concat'),
    uglify = require('gulp-uglify'),
    rename = require('gulp-rename'),
    jshint = require('gulp-jshint');

gulp.task('minifycss', function () {
    return gulp.src('src/Office.Controls.PeoplePicker.css')
        .pipe(rename({ suffix: '.min' }))
        .pipe(minifycss())
        .pipe(gulp.dest('dist/'));
});

gulp.task('minifyjs', function () {
    return gulp.src('src/Office.Controls.PeoplePicker.js')
        .pipe(rename({ suffix: '.min' }))
        .pipe(uglify({compress: true,mangle: true, outSourceMap: true}))
        .pipe(gulp.dest('dist/'));
});

gulp.task('runjshint', function () {
    return gulp.src('src/Office.Controls.PeoplePicker.js')
      .pipe(jshint('tools/jshint/.jshintrc.json'))
      .pipe(jshint.reporter('jshint-stylish'));
});

gulp.task('cpfiles', function() {
    ["src/Office.Controls.PeoplePicker.css",
    "src/Office.Controls.PeoplePicker.js"
    ].forEach(
        function (file) {
            gulp.src(file)
            .pipe(gulp.dest('dist/'));
        });

    gulp.src("src/css/*")
    .pipe(gulp.dest('dist/css/'));
    
    gulp.src("dist/*")
    .pipe(gulp.dest('example/control/'));
});

gulp.task('default', function () {
    gulp.start('runjshint', 'cpfiles', 'minifycss', 'minifyjs');
});