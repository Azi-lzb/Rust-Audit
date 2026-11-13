# Included verbatim in deb maintainer scripts; essential POSIX shell/coreutils only.
# Active configuration is deliberately NOT package-owned. New defaults live separately.
audit_config_snapshot() {
    active=$1
    saved=$2
    if [ -d "$active" ]; then
        mkdir -p "$saved"
        chmod 700 "$(dirname "$saved")"
        cp -a "$active"/. "$saved"/
    fi
}

audit_config_restore() {
    active=$1
    saved=$2
    defaults=$3
    mkdir -p "$active"
    # cp -n never replaces existing user files, including when configure is retried.
    if [ -d "$saved" ]; then cp -a -n "$saved"/. "$active"/; fi
    if [ -d "$defaults" ]; then cp -a -n "$defaults"/. "$active"/; fi
}
