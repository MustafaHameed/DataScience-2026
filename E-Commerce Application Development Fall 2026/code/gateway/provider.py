"""A payment provider that does not exist, so that you can measure one.

Chapter 14 needs a gateway. Integrating a real one needs a merchant account,
which needs a registered business, which a student does not have -- and a
sandbox account needs credentials that cannot be put in a handout and will have
expired by the time anyone reads it. So the course ships its own provider.

This is not a toy version of a gateway. It is a gateway reduced to the three
properties that make integrating with one difficult:

  1. IT IS SLOW, AND IRREGULARLY SO. A real provider answers in a few hundred
     milliseconds, usually. The distribution has a long tail, and the tail is
     what your timeout has to be chosen against.
  2. IT SOMETIMES DOES NOT ANSWER AT ALL. A request that times out has an
     unknown outcome: the payment may have been taken.
  3. IT TELLS YOU TWICE. Providers retry their callbacks when they are not
     acknowledged quickly enough, so your shop WILL receive the same
     notification more than once, and must not act on it more than once.

Everything is configurable from the environment so that a student can make it
behave badly on purpose:

    GATEWAY_DELAY_MS     median response delay            (default 300)
    GATEWAY_JITTER_MS    uniform jitter either side       (default 200)
    GATEWAY_TAIL_RATE    share of requests that are slow  (default 0.10)
    GATEWAY_TAIL_MS      how slow those are               (default 2500)
    GATEWAY_FAIL_RATE    share that decline               (default 0.05)
    GATEWAY_HANG_RATE    share that never answer          (default 0.02)

Endpoints:

    GET  /health                  is it up
    POST /authorize               {order_id, amount, idempotency_key?}
    GET  /stats                   what it has seen, for the harness
    POST /reset                   forget everything

It is deliberately a single file with no dependencies beyond the standard
library, so that it runs in a stock python image with nothing installed and
a student can read the whole of it in ten minutes.
"""
import json
import os
import random
import threading
import time
import uuid
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer

PORT = int(os.environ.get("GATEWAY_PORT", "8092"))
DELAY = int(os.environ.get("GATEWAY_DELAY_MS", "300"))
JITTER = int(os.environ.get("GATEWAY_JITTER_MS", "200"))
TAIL_RATE = float(os.environ.get("GATEWAY_TAIL_RATE", "0.10"))
TAIL_MS = int(os.environ.get("GATEWAY_TAIL_MS", "2500"))
FAIL_RATE = float(os.environ.get("GATEWAY_FAIL_RATE", "0.05"))
HANG_RATE = float(os.environ.get("GATEWAY_HANG_RATE", "0.02"))

LOCK = threading.Lock()
SEEN: dict[str, dict] = {}          # idempotency key -> the first answer
COUNTS = {"authorize": 0, "approved": 0, "declined": 0,
          "hung": 0, "replayed": 0}


def hold() -> float:
    """How long this request will take, in seconds, and whether it hangs.

    A real provider's latency is not a bell curve around its median. Most
    answers are quick and a minority are very slow, because a minority go
    through an extra hop -- a risk check, an issuer that is busy, a retry
    inside the provider's own network. Modelling it as median plus jitter
    alone would produce a timeout recommendation that is far too low.
    """
    if random.random() < TAIL_RATE:
        return (TAIL_MS + random.uniform(0, TAIL_MS)) / 1000.0
    base = DELAY + random.uniform(-JITTER, JITTER)
    return max(0.01, base / 1000.0)


class Handler(BaseHTTPRequestHandler):
    protocol_version = "HTTP/1.1"

    def log_message(self, *args):          # keep the container log readable
        pass

    def reply(self, code: int, body: dict) -> None:
        raw = json.dumps(body).encode()
        self.send_response(code)
        self.send_header("Content-Type", "application/json")
        self.send_header("Content-Length", str(len(raw)))
        self.end_headers()
        self.wfile.write(raw)

    def do_GET(self):
        if self.path == "/health":
            self.reply(200, {"status": "up", "delay_ms": DELAY,
                             "tail_rate": TAIL_RATE, "fail_rate": FAIL_RATE,
                             "hang_rate": HANG_RATE})
        elif self.path == "/stats":
            with LOCK:
                self.reply(200, dict(COUNTS, keys=len(SEEN)))
        else:
            self.reply(404, {"error": "no such endpoint"})

    def do_POST(self):
        if self.path == "/reset":
            with LOCK:
                SEEN.clear()
                for k in COUNTS:
                    COUNTS[k] = 0
            self.reply(200, {"status": "reset"})
            return
        if self.path != "/authorize":
            self.reply(404, {"error": "no such endpoint"})
            return

        n = int(self.headers.get("Content-Length", "0"))
        try:
            req = json.loads(self.rfile.read(n) or b"{}")
        except ValueError:
            self.reply(400, {"error": "bad json"})
            return

        key = req.get("idempotency_key")
        with LOCK:
            COUNTS["authorize"] += 1
            if key and key in SEEN:
                # The whole point of an idempotency key: the second request
                # is answered from the first answer, without charging again.
                COUNTS["replayed"] += 1
                prior = dict(SEEN[key], replayed=True)
                self.reply(200, prior)
                return

        wait = hold()
        if random.random() < HANG_RATE:
            with LOCK:
                COUNTS["hung"] += 1
            time.sleep(30)             # longer than any sane client waits
            return

        time.sleep(wait)
        approved = random.random() >= FAIL_RATE
        answer = {
            "order_id": req.get("order_id"),
            "amount": req.get("amount"),
            "status": "approved" if approved else "declined",
            "transaction_id": uuid.uuid4().hex[:16],
            "took_ms": round(wait * 1000),
            "replayed": False,
        }
        with LOCK:
            COUNTS["approved" if approved else "declined"] += 1
            if key:
                SEEN[key] = answer
        self.reply(200, answer)


if __name__ == "__main__":
    random.seed()
    print(f"mock gateway on :{PORT}  delay={DELAY}ms jitter={JITTER}ms "
          f"tail={TAIL_RATE} fail={FAIL_RATE} hang={HANG_RATE}", flush=True)
    ThreadingHTTPServer(("0.0.0.0", PORT), Handler).serve_forever()
