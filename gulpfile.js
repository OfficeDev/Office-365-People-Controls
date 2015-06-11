var gulp = require('gulp'),
    minifycss = require('gulp-minify-css'),
    concat = require('gulp-concat'),
    uglify = require('gulp-uglify'),
    rename = require('gulp-rename'),
    jshint = require('gulp-jshint');

gulp.task('minifycss', function () {
    return ["src/Office.Controls.PeoplePicker.css",
            "src/Office.Controls.Persona.css"
           ].forEach(
                function (file) {
                    gulp.src(file)
                    .pipe(rename({ suffix: '.min' }))
                    .pipe(minifycss())
                    .pipe(gulp.dest('dist/'));
                });
});

gulp.task('minifyjs', function () {
    return ["src/Office.Controls.Base.js",
            "src/Office.Controls.PeopleAadDataProvider.js",
            "src/Office.Controls.PeoplePicker.js",
            "src/Office.Controls.Persona.js"
           ].forEach(
                function (file) {
                    gulp.src(file)
                    .pipe(rename({ suffix: '.min' }))
                    .pipe(uglify({compress: true,mangle: true, outSourceMap: true}))
                    .pipe(gulp.dest('dist/'));
                });
});

gulp.task('runjshint', function () {
    return ["src/Office.Controls.Base.js",
            "src/Office.Controls.PeopleAadDataProvider.js",
            "src/Office.Controls.PeoplePicker.js",
            "src/Office.Controls.Persona.js"
           ].forEach(
                function (file) {
                    gulp.src(file)
                    .pipe(jshint('tools/jshint/.jshintrc.json'))
                    .pipe(jshint.reporter('jshint-stylish'));
                });
});

gulp.task('cpfiles', function() {
    ["src/Office.Controls.Base.js",
     "src/Office.Controls.PeopleAadDataProvider.js",
     "src/Office.Controls.PeoplePicker.css",
     "src/Office.Controls.PeoplePicker.js",
     "src/Office.Controls.Persona.css",
     "src/Office.Controls.Persona.js"
    ].forEach(
        function (file) {
            gulp.src(file)
            .pipe(gulp.dest('dist/'));
        });

    gulp.src("src/css/*")
    .pipe(gulp.dest('dist/css/'));

    gulp.src("src/templates/*")
    .pipe(gulp.dest('dist/templates/'));
    
    gulp.src("dist/**")
    .pipe(gulp.dest('example/control/'));
});

gulp.task('default', function () {
    gulp.start('runjshint', 'minifycss', 'minifyjs', 'cpfiles');
});