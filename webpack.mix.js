const mix = require('laravel-mix')

const autoload = {
   jquery: [ '$', 'jQuery', 'jquery']
}
mix.autoload(autoload)

mix.setPublicPath('public')

mix.js('./resources/js/app.js', 'public/js')
   .sass('./resources/sass/app.scss', 'public/css')
   .version()

// Copy all compiled files into main project (auto publishing)
mix.copyDirectory('public', '../../../public/vendor/uccello/document-designer-core');