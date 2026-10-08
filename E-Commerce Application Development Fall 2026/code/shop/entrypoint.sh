#!/bin/sh
# Install OpenCart on first start, then hand over to Apache.
#
# The install cannot happen at image-build time, because it needs a database
# that is running -- and nothing is running during a build.  So it happens
# here, once, guarded by a lock file that lives in the same volume as the
# install itself.  "docker compose down -v" removes the volume and the next
# "up" installs again from scratch, which is how a student resets a shop they
# have broken, and how run_all.py guarantees every measurement starts from the
# same state.
set -eu

HTDOCS=/var/www/html
LOCK="$HTDOCS/install.lock"

if [ ! -f "$LOCK" ]; then
  echo "[entrypoint] no install.lock -- installing OpenCart"

  # Wait for the database.  depends_on only orders container *starts*; MariaDB
  # accepts TCP connections several seconds before it will accept a login, so
  # ordering alone is not enough and the install would fail on a fast machine.
  i=0
  until php -r 'exit(@mysqli_connect(getenv("DB_HOST"), getenv("DB_USER"), getenv("DB_PASSWORD"), getenv("DB_NAME")) ? 0 : 1);'; do
    i=$((i + 1))
    if [ "$i" -ge 60 ]; then
      echo "[entrypoint] database did not accept a login after 60 tries" >&2
      exit 1
    fi
    echo "[entrypoint] waiting for database ($i)"
    sleep 2
  done

  php "$HTDOCS/install/cli_install.php" install \
      --username    "$OC_ADMIN_USER" \
      --password    "$OC_ADMIN_PASSWORD" \
      --email       "$OC_ADMIN_EMAIL" \
      --http_server "$OC_HTTP_SERVER" \
      --db_driver   mysqli \
      --db_hostname "$DB_HOST" \
      --db_username "$DB_USER" \
      --db_password "$DB_PASSWORD" \
      --db_database "$DB_NAME" \
      --db_port     3306 \
      --db_prefix   oc_

  # OpenCart's own INSTALL.md step 8: remove the installer once it has run.
  # Leaving it in place lets anyone who can reach the site reinstall the shop
  # over the top of yours.  Chapter 16 measures what else is left behind.
  rm -rf "$HTDOCS/install"

  chown -R www-data:www-data "$HTDOCS"
  touch "$LOCK"
  echo "[entrypoint] install complete"
else
  echo "[entrypoint] install.lock present -- starting without installing"
fi

exec "$@"
