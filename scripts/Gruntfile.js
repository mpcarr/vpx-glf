module.exports = function (grunt) {
  grunt.initConfig({
    concat: {
      vpx: {
        src: ['src/**/*.vbs', '!src/unittests/**/*.vbs', '!src/**/*.test.vbs'],
        dest: './vpx-glf.vbs',
      },
    },
    watch: {
      vpx: {
        files: 'src/**/*.vbs',
        tasks: ['concat:vpx'],
        options: {
          atBegin: true
        }
      },
    }

  });

  grunt.loadNpmTasks('grunt-contrib-concat');
  grunt.loadNpmTasks('grunt-contrib-watch');
  //grunt.loadNpmTasks('grunt-exec');


  // Default task(s).
  grunt.registerTask('default', ['concat']);

};