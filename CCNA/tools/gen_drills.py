#!/usr/bin/env python3
"""Generate the drill tables for Appendix C, with answers computed rather
than typed.

Subnetting answers in a printed reference are exactly the place a
transcription error survives unnoticed -- a reader who gets a different
answer assumes they are wrong. Every value here comes from Python's
ipaddress module, so the book cannot be wrong unless the standard library
is.

    python tools/gen_drills.py > drills.tex

Then paste the two tables into parts/C_subnetting.tex.
"""
import ipaddress

IPV4_DRILLS = [
    "192.168.10.77/26",   "10.4.19.200/21",     "172.16.35.14/20",
    "192.168.1.130/25",   "10.0.0.66/30",       "172.20.140.9/22",
    "192.168.200.35/27",  "10.55.12.3/19",      "172.31.255.200/18",
    "192.168.16.62/28",   "10.128.64.31/27",    "172.16.0.129/29",
    "203.0.113.45/29",    "198.51.100.17/28",   "192.0.2.200/26",
    "10.10.10.10/23",     "172.18.6.250/21",    "192.168.99.99/24",
]

HOST_REQUIREMENTS = [2, 6, 12, 25, 50, 100, 200, 500, 1000]

IPV6_COMPRESS = [
    "2001:0db8:0000:0000:0000:ff00:0042:8329",
    "2001:0db8:acad:0010:0000:0000:0000:0001",
    "fe80:0000:0000:0000:0204:61ff:fe9d:f156",
    "2001:0db8:0000:0001:0000:0000:0000:0100",
    "0000:0000:0000:0000:0000:0000:0000:0001",
    "2001:0db8:aaaa:0001:0000:0000:0000:0000",
]

EUI64 = [
    ("2001:db8:acad:10::/64", "00:1a:2b:3c:4d:5e"),
    ("2001:db8:1:1::/64",     "aa:bb:cc:dd:ee:ff"),
    ("fe80::/64",             "0c:1d:2e:3f:40:51"),
    ("2001:db8:99::/64",      "52:54:00:12:34:56"),
]


def row(shade, cells):
    out = []
    if shade:
        out.append("\\rowcolor{lightbg}")
    out.append(" & ".join(cells) + " \\\\")
    return out


def ipv4_table():
    # Column budgets for this text block: 15.78 cm for two columns, 15.36
    # for three, 14.93 for four, 14.51 for five. A dotted-quad in \cmd is
    # about 0.185 cm per character and cannot be broken, so the column has
    # to fit the longest value outright.
    L = ["\\begin{longtable}{@{}L{0.8cm}L{4.2cm}L{3.1cm}L{3.1cm}L{3.3cm}@{}}",
         "\\toprule", "\\rowcolor{primary!15}",
         "\\textbf{} & \\textbf{Given} & \\textbf{Network} & "
         "\\textbf{Broadcast} & \\textbf{Usable hosts} \\\\",
         "\\midrule", "\\endhead"]
    ans = []
    for i, spec in enumerate(IPV4_DRILLS, 1):
        iface = ipaddress.ip_interface(spec)
        net = iface.network
        hosts = net.num_addresses - 2 if net.prefixlen < 31 else 0
        first = net.network_address + 1
        last = net.broadcast_address - 1
        L += row(i % 2 == 0, [str(i), "\\cmd{%s}" % spec, "", "", ""])
        ans.append((i, spec, str(net.network_address), str(net.broadcast_address),
                    str(first), str(last), hosts, str(net.netmask)))
    L += ["\\bottomrule", "\\end{longtable}"]

    # The usable range is two addresses, which will not fit on one line in
    # any column this table can afford, so it is stacked.
    A = ["\\begin{longtable}{@{}L{0.8cm}L{3.1cm}L{3.0cm}L{3.0cm}L{4.4cm}@{}}",
         "\\toprule", "\\rowcolor{primary!15}",
         "\\textbf{} & \\textbf{Mask} & \\textbf{Network} & "
         "\\textbf{Broadcast} & \\textbf{Usable range (count)} \\\\",
         "\\midrule", "\\endhead"]
    for i, spec, net, bc, first, last, hosts, mask in ans:
        A += row(i % 2 == 0, [
            str(i), "\\cmd{%s}" % mask, "\\cmd{%s}" % net, "\\cmd{%s}" % bc,
            "\\cmd{%s} \\newline \\cmd{to %s} (%d)" % (first, last, hosts)])
    A += ["\\bottomrule", "\\end{longtable}"]
    return L, A


def sizing_table():
    L = ["\\begin{longtable}{@{}L{3.4cm}L{3.0cm}L{3.4cm}L{4.6cm}@{}}",
         "\\toprule", "\\rowcolor{primary!15}",
         "\\textbf{Hosts needed} & \\textbf{Prefix} & \\textbf{Mask} & "
         "\\textbf{Usable / wasted} \\\\",
         "\\midrule", "\\endhead"]
    for i, need in enumerate(HOST_REQUIREMENTS, 1):
        bits = 2
        while (2 ** bits) - 2 < need:
            bits += 1
        prefix = 32 - bits
        net = ipaddress.ip_network("10.0.0.0/%d" % prefix)
        usable = net.num_addresses - 2
        L += row(i % 2 == 0, [
            str(need), "\\cmd{/%d}" % prefix, "\\cmd{%s}" % net.netmask,
            "%d / %d" % (usable, usable - need)])
    L += ["\\bottomrule", "\\end{longtable}"]
    return L


def ipv6_tables():
    Q = ["\\begin{longtable}{@{}L{0.8cm}L{7.6cm}L{6.9cm}@{}}",
         "\\toprule", "\\rowcolor{primary!15}",
         "\\textbf{} & \\textbf{Given} & \\textbf{Compressed} \\\\",
         "\\midrule", "\\endhead"]
    A = list(Q)
    for i, addr in enumerate(IPV6_COMPRESS, 1):
        short = str(ipaddress.IPv6Address(addr))
        Q += row(i % 2 == 0, [str(i), "\\cmd{%s}" % addr, ""])
        A += row(i % 2 == 0, [str(i), "\\cmd{%s}" % addr, "\\cmd{%s}" % short])
    Q += ["\\bottomrule", "\\end{longtable}"]
    A += ["\\bottomrule", "\\end{longtable}"]

    E = ["\\begin{longtable}{@{}L{0.7cm}L{4.0cm}L{3.2cm}L{7.0cm}@{}}",
         "\\toprule", "\\rowcolor{primary!15}",
         "\\textbf{} & \\textbf{Prefix} & \\textbf{MAC} & "
         "\\textbf{EUI-64 address} \\\\",
         "\\midrule", "\\endhead"]
    EA = list(E)
    for i, (prefix, mac) in enumerate(EUI64, 1):
        h = mac.replace(":", "")
        first = int(h[0:2], 16) ^ 0x02          # flip the U/L bit
        eui = "%02x%s:%sff:fe%s:%s" % (first, h[2:4], h[4:6], h[6:8], h[8:12])
        net = ipaddress.ip_network(prefix)
        full = ipaddress.IPv6Address(
            int(net.network_address) | int(ipaddress.IPv6Address("::" + eui)))
        E += row(i % 2 == 0, [str(i), "\\cmd{%s}" % prefix, "\\cmd{%s}" % mac, ""])
        EA += row(i % 2 == 0, [str(i), "\\cmd{%s}" % prefix, "\\cmd{%s}" % mac,
                               "\\cmd{%s}" % full])
    E += ["\\bottomrule", "\\end{longtable}"]
    EA += ["\\bottomrule", "\\end{longtable}"]
    return Q, A, E, EA


if __name__ == "__main__":
    q4, a4 = ipv4_table()
    sz = sizing_table()
    q6, a6, qe, ae = ipv6_tables()
    for name, block in [("IPV4 QUESTIONS", q4), ("IPV4 ANSWERS", a4),
                        ("SIZING (answers shown)", sz),
                        ("IPV6 COMPRESS Q", q6), ("IPV6 COMPRESS A", a6),
                        ("EUI64 Q", qe), ("EUI64 A", ae)]:
        print("\n%% ======== %s ========" % name)
        print("\n".join(block))
