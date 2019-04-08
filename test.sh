#!/bin/bash

docker run -it --rm -v $(pwd):/var/www/html cigzigwon/php:7.1-fpm-alpine phpunit --bootstrap vendor/autoload.php tests/