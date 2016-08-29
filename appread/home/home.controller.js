(function(){
  'use strict';

  angular.module('officeAddin')
         .controller('homeController', [homeController]);

  /**
   * Controller constructor
   */
  function homeController(){
    var vm = this;  // jshint ignore:line
    vm.title = 'home controller';
    vm.dataObject = {};
  }

})();
