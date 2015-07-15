var gulp = require('gulp'),
    minifycss = require('gulp-minify-css'),
    concat = require('gulp-concat'),
    uglify = require('gulp-uglify'),
    rename = require('gulp-rename'),
    jshint = require('gulp-jshint');

gulp.task('minifycss', ['concatcss'], function () {
    return ['src/Office.Controls.PeoplePicker.css',
            'src/Office.Controls.Persona.css',
            'dist/Office.Controls.People.css'
           ].forEach(
                function (file) {
                    gulp.src(file)
                    .pipe(rename({ suffix: '.min' }))
                    .pipe(minifycss())
                    .pipe(gulp.dest('dist/'));
                });
});

gulp.task('minifyjs', ['concatjs'], function () {
    return ['src/Office.Controls.Base.js',
            'src/Office.Controls.PeopleAadDataProvider.js',
            'src/Office.Controls.PeoplePicker.js',
            'src/Office.Controls.Persona.js',
            'dist/Office.Controls.People.js'
           ].forEach(
                function (file) {
                    gulp.src(file)
                    .pipe(rename({ suffix: '.min' }))
                    .pipe(uglify({compress: true,mangle: true, outSourceMap: true}))
                    .pipe(gulp.dest('dist/'));
                });
});

gulp.task('runjshint', function () {
    return ['src/Office.Controls.Base.js',
            'src/Office.Controls.PeopleAadDataProvider.js',
            'src/Office.Controls.PeoplePicker.js',
            'src/Office.Controls.Persona.js'
           ].forEach(
                function (file) {
                    gulp.src(file)
                    .pipe(jshint('tools/jshint/.jshintrc.json'))
                    .pipe(jshint.reporter('jshint-stylish'));
                });
});

gulp.task('concatjs', function() {
    return gulp.src(['src/Office.Controls.Base.js', 'src/Office.Controls.PeopleAadDataProvider.js', 'src/Office.Controls.PeoplePicker.js', 'src/Office.Controls.Persona.js'])
    .pipe(concat('Office.Controls.People.js'))
    .pipe(gulp.dest('dist/'));
});

gulp.task('concatcss', function() {
    return gulp.src(['src/Office.Controls.PeoplePicker.css', 'src/Office.Controls.Persona.css'])
    .pipe(concat('Office.Controls.People.css'))
    .pipe(gulp.dest('dist/'));
});

gulp.task('cpfilestodist', ['minifycss', 'minifyjs'], function() {
    return gulp.src('src/**/*')
    .pipe(gulp.dest('dist/'));
});

gulp.task('cpfilestoexample', ['cpfilestodist'], function() {
    return gulp.src('dist/**/*')
    .pipe(gulp.dest('example/control/'));
});

gulp.task('default', ['runjshint', 'cpfilestoexample']);