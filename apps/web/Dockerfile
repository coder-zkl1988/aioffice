ARG ALPINE_MIRROR

FROM rust:alpine AS xlsx-sidecar-builder

ARG ALPINE_MIRROR
ARG CARGO_REGISTRY

RUN if [ -n "$ALPINE_MIRROR" ]; then \
      sed -i "s#dl-cdn.alpinelinux.org#$ALPINE_MIRROR#g" /etc/apk/repositories; \
    fi \
 && apk add --no-cache musl-dev

WORKDIR /build
COPY apps/sheets/native/xlsx-engine/Cargo.toml apps/sheets/native/xlsx-engine/Cargo.lock ./
COPY apps/sheets/native/xlsx-engine/src ./src
RUN if [ -n "$CARGO_REGISTRY" ]; then \
      cargo_config_dir="${CARGO_HOME:-/root/.cargo}"; \
      mkdir -p "$cargo_config_dir"; \
      printf '[source.crates-io]\nreplace-with = "mirror"\n\n[source.mirror]\nregistry = "%s"\n' \
        "$CARGO_REGISTRY" > "$cargo_config_dir/config.toml"; \
    fi \
 && cargo build --release --locked

FROM node:24-alpine AS web-builder

ARG NPM_REGISTRY

WORKDIR /source
COPY . .
RUN if [ -n "$NPM_REGISTRY" ]; then npm config set registry "$NPM_REGISTRY"; fi \
 && npm ci --ignore-scripts
COPY --from=xlsx-sidecar-builder /build/target/release/xlsx-sidecar \
  /source/apps/sheets/native/xlsx-engine/target/release/xlsx-sidecar
RUN npm run build:web

FROM node:24-alpine

ARG ALPINE_MIRROR
RUN if [ -n "$ALPINE_MIRROR" ]; then \
      sed -i "s#dl-cdn.alpinelinux.org#$ALPINE_MIRROR#g" /etc/apk/repositories; \
    fi \
 && apk add --no-cache font-dejavu font-liberation font-noto-cjk libgcc

WORKDIR /app
COPY --from=web-builder /source/apps/web/dist ./dist
COPY --from=web-builder /source/apps/web/dist-server ./dist-server

ENV GENOFFICE_WASM_DIR=/app/dist-server/wasm
ENV XLSX_SIDECAR_PATH=/app/dist-server/native/xlsx-sidecar

EXPOSE 80

CMD ["node", "dist-server/server.mjs"]
