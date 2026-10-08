# The storefront layer of the teaching stack.
#
# Built from the official PHP image rather than pulled ready-made, for two
# reasons.  The first is that there is no longer a maintained OpenCart image
# to pull: Bitnami withdrew theirs and there has never been an official one,
# so every "docker pull opencart" instruction on the web is now dead.  The
# second is that Chapter 5 is about what a stack is made of, and a student
# who pulls one opaque image has not seen a stack -- they have seen a magic
# word.  Here the layers are visible: Debian, Apache, PHP, the extensions
# OpenCart needs, then the application.
#
# OpenCart ships its own Dockerfile in the release zip.  It is not used here
# and Chapter 20 explains why: it contains a shell syntax error ("&& &&") and
# so has never built, and its companion compose file pins mysql:5.7, which
# left support in October 2023.  Upstream artefacts are evidence, not
# authority.

FROM php:8.2-apache

# The release is pinned by version AND by checksum.  A pinned version alone
# is not reproducible: a tag can be moved, and a release asset can be
# replaced.  If this build ever fails on the checksum, do not "fix" it by
# deleting the check -- find out what changed.
ARG OC_VERSION=4.1.0.4
ARG OC_SHA256=cdef537f5a0a7e1b4ddbe16f589dd2e80b10150a09bfbe93ce76ef2d5c3218c0

# gd, zip and mysqli are OpenCart's hard requirements; opcache is not, but
# without it every page recompiles every PHP file on every request and the
# Chapter 20 measurements are a study of the compiler rather than the shop.
RUN set -eux; \
    apt-get update; \
    apt-get install -y --no-install-recommends \
        unzip libfreetype6-dev libjpeg62-turbo-dev libpng-dev libzip-dev; \
    docker-php-ext-configure gd --with-freetype --with-jpeg; \
    docker-php-ext-install -j"$(nproc)" gd zip mysqli opcache; \
    a2enmod rewrite; \
    rm -rf /var/lib/apt/lists/*

# Without this Apache writes "AH00558: Could not reliably determine the
# server's fully qualified domain name" into the log twice on every start.
# It is harmless, and it is also the first thing a student sees when they go
# looking for a real error, which is reason enough to silence it.
RUN printf 'ServerName localhost\n' > /etc/apache2/conf-available/servername.conf \
    && a2enconf servername

RUN set -eux; \
    curl -fsSL -o /tmp/oc.zip \
      "https://github.com/opencart/opencart/releases/download/${OC_VERSION}/opencart-${OC_VERSION}.zip"; \
    echo "${OC_SHA256}  /tmp/oc.zip" | sha256sum -c -; \
    unzip -q /tmp/oc.zip -d /tmp/oc; \
    cp -a /tmp/oc/upload/. /var/www/html/; \
    rm -rf /tmp/oc.zip /tmp/oc

# OpenCart ships config-dist.php and expects you to copy it.  Doing that here
# rather than at every start keeps the entrypoint to one job.
RUN set -eux; \
    cp /var/www/html/config-dist.php       /var/www/html/config.php; \
    cp /var/www/html/admin/config-dist.php /var/www/html/admin/config.php; \
    chown -R www-data:www-data /var/www/html

COPY php.ini /usr/local/etc/php/conf.d/zz-shop.ini
COPY entrypoint.sh /usr/local/bin/entrypoint.sh
RUN chmod +x /usr/local/bin/entrypoint.sh

# The storage directory stays inside the web root, which is NOT where it
# belongs.  That is deliberate: it is the stock install, it is what a student
# following the vendor's own instructions will have, and moving it out is one
# of the hardening steps measured in Chapter 16.  A course that hands over an
# already-hardened install has nothing left to teach about hardening.

EXPOSE 80
ENTRYPOINT ["/usr/local/bin/entrypoint.sh"]
CMD ["apache2-foreground"]
