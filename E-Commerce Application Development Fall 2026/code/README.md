# `code/` — the stack, the harnesses and the modules

Everything the handout measures was measured by something in this folder, and
every figure it quotes can be reproduced by running the named file. This README
says what each file is for, what it changes, and what it will not do.

## An honest deviation from the sibling courses

The other six courses in this repository hold about twenty-three files in one
language, one per chapter, each a self-contained program. A platform course
cannot be arranged that way and pretending otherwise would be worse than saying
so: this folder holds **a stack definition, a mock payment provider, twelve
measurement harnesses, a catalogue seeder, two Apache configurations and a PHP
extension**, in five languages. The structure follows the subject.

## Start here

```bash
docker compose up -d --build     # first run builds and installs; a few minutes
python waitfor.py                # all three URLs answer, or it says which did not
```

| | URL | What it is |
|---|---|---|
| Storefront | <http://localhost:8090/> | The shop |
| Admin | <http://localhost:8090/admin/> | `admin` / `admin123` |
| Adminer | <http://localhost:8091/> | The database, directly |
| Gateway | <http://localhost:8092/health> | The mock payment provider (Ch 14) |

`docker compose down` keeps your shop. **`docker compose down -v` throws it
away** — Chapter 9 measures exactly that difference, and it takes 17 seconds to
rebuild from nothing.

The credentials above are weak on purpose and documented rather than hidden:
they belong to a container on your own machine holding no real person's data.
Chapter 16 and Chapter 20 both say what must change before any shop leaves it.

## The stack

| File | What it is |
|---|---|
| `docker-compose.yml` | Four services: `db` (MariaDB 11), `shop`, `adminer`, `gateway`. No `version:` key — obsolete since Compose v2 |
| `shop/Dockerfile` | `php:8.2-apache` + the extensions OpenCart needs. The release is pinned **by version and by SHA-256** |
| `shop/entrypoint.sh` | Waits for the database to accept a login — not just the port — then installs OpenCart and removes `install/` |
| `shop/php.ini` | Upload limits, opcache |
| `shop/security.conf` | Chapter 16's hardening. **Not enabled by default**; the chapter measures the shop without it, then with it |
| `shop/seourl.conf` | Chapter 21's rewrite — the server half of readable URLs |
| `gateway/provider.py` | A payment provider that does not exist. Fifty lines of standard library, slow and irregular and occasionally silent *on purpose* |
| `waitfor.py` | Polls all three URLs; used by every harness so none of them measures a shop that is still starting |

## The measurement harnesses

Each is named for the chapter that quotes it. Run with `--help` for its modes.

| File | Chapter | What it measures | Changes the shop? |
|---|---|---|---|
| `m05_coldstart.py` | 5 | Cold, warm and from-scratch start; stack footprint | **Yes — `--warm` avoids it** |
| `m06_catalogue.py` | 6 | Page time and queries against catalogue size, with and without indexes | **Yes — seeds products** |
| `m07_variations.py` | 7 | Options against variants: rows, combinations, page impact | **Yes** |
| `m08_pageweight.py` | 8 | Bytes and requests per page, split by kind | No |
| `m09_basket.py` | 9 | Where basket state lives; session engines; what survives a restart | Yes (cleans up) |
| `m10_checkout.py` | 10 | Per-step time, queries and **writes**; the order state machine | Yes — `--clean` |
| `m12_discounts.py` | 12 | A coupon before and after the tax line | Yes — `--restore` |
| `m14_gateway.py` | 14 | Gateway latency distribution, timeouts, idempotency | No |
| `m15_accounts.py` | 15 | bcrypt cost against logins per second; user enumeration; lockout | Yes — `--clean` |
| `m16_security.py` | 16 | Fifteen checks, before and after hardening | Yes — `--unharden` |
| `m17_dashboard.py` | 17 | What each part of the dashboard counts | Yes — `--clean` |
| `m18_orders.py` | 18 | The admin API; which status changes move stock | Yes — `--clean` |
| `m19_modules.py` | 19 | What a module costs a page, at 0, 1, 3 and 5 instances | Yes — `--remove` |
| `m20_deploy.py` | 20 | Throughput and latency under load; opcache; backup and restore | Restarts Apache |
| `m21_seo.py` | 21 | Readable URLs, titles, descriptions, canonical, sitemap, first byte | Yes — `--restore` |

```bash
python run_all.py --quick     # every harness except the three slow ones
python run_all.py --list      # what it would run, and what it checks
python sync_labs.py           # every command the handout prints actually exists
```

`run_all.py` does not check that the numbers are the same — they are facts
about one laptop. It checks that the **shape** of each result still holds: that
the session record stays constant while the basket grows, that the dashboard's
three counts still disagree, that there is still no sitemap. If one fails, the
handout and the software have diverged, and the first job is to find out which
one moved.

## The seeder

```bash
python seed/catalogue.py 50000     # reproducible, chunked, ~2 minutes
python seed/catalogue.py --restore # back to the 19 demo products
```

`STORE = 0`, and that constant has a scar on it. An earlier version wrote
`store_id = 1`; the default store is 0, so fifty thousand products were
invisible to every customer while rendering perfectly in the admin panel.
Chapter 6 had to be measured again from scratch. Every page returned 200
throughout.

## The extension

`modules/ecadmod/` is the extension written in Chapter 19 — twelve files
providing a storefront module (`lowstock`) and a shipping method (`zonerate`).

```bash
python m19_modules.py --install    # files, manifest row, one row per thing
python m19_modules.py --remove     # all of it, cleanly
```

It touches nothing OpenCart ships. That is the whole point of the chapter, and
`--remove` proving the shop is exactly as it was is how you know.

## The machine every timing was taken on

Appendix A of the handout states it. Every number in this folder's output is a
fact about that machine; the comparisons between configurations are what
transfer, not the absolute figures.

## What will not work here

- **The host has no PHP and no MySQL**, and does not need them. Everything runs
  in the containers.
- **Ports 8080 and 8081 were already in use** on the machine this was written
  on, which is why the defaults are 8090–8092. `SHOP_PORT`, `ADMINER_PORT` and
  `GATEWAY_PORT` override them.
- **`docker compose exec db mysql`** — the MariaDB 11 image ships `mariadb`,
  not `mysql`.
- **Tailing `/var/log/apache2/access.log` inside the container hangs**: it is a
  symlink to `/dev/stdout`. Use `docker compose logs shop`.
