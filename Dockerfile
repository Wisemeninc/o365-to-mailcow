# syntax=docker/dockerfile:1
# o365-to-mailcow: official Python 3.12 slim image, pinned by tag and digest.
ARG PYTHON_IMAGE=python:3.12.14-slim-bookworm@sha256:392307d22300de8b5986851a12d9176dfc0fc073e65bf6523ebd7dcbeb23564e

# --- build: turn the project into a wheel (build tooling never reaches the final image)
FROM ${PYTHON_IMAGE} AS build
ENV PIP_NO_CACHE_DIR=1 PIP_DISABLE_PIP_VERSION_CHECK=1
WORKDIR /src
COPY pyproject.toml README.md ./
COPY src ./src
RUN pip wheel --no-deps --wheel-dir /wheels .

# --- runtime
FROM ${PYTHON_IMAGE}
LABEL org.opencontainers.image.title="o365-to-mailcow" \
      org.opencontainers.image.description="Migrate Microsoft 365 mail, calendars and contacts into mailcow via Microsoft Graph" \
      org.opencontainers.image.licenses="MIT"
ENV PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1 \
    PIP_NO_CACHE_DIR=1 \
    PIP_DISABLE_PIP_VERSION_CHECK=1 \
    O365MIG_CONFIG=/config/config.toml

# Dependencies from the hash-pinned lock file, then the project wheel without resolving.
COPY requirements.txt /tmp/requirements.txt
RUN pip install --require-hashes --only-binary=:all: -r /tmp/requirements.txt \
    && rm /tmp/requirements.txt
COPY --from=build /wheels /tmp/wheels
RUN pip install --no-deps /tmp/wheels/*.whl && rm -rf /tmp/wheels

RUN groupadd --system --gid 10001 o365mig \
    && useradd --system --uid 10001 --gid o365mig --home-dir /state \
       --shell /usr/sbin/nologin o365mig \
    && mkdir -p /state /config \
    && chown o365mig:o365mig /state \
    && chmod 0700 /state

USER o365mig
WORKDIR /state
VOLUME /state
ENTRYPOINT ["o365mig"]
CMD ["--help"]
