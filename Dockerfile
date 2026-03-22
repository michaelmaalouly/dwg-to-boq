FROM python:3.12-slim

# Install build tools, compile LibreDWG, then clean up
# We only need dwg2dxf (read-only), so build just the programs target
# and suppress all -Werror flags that break on newer GCC
RUN apt-get update && apt-get install -y --no-install-recommends \
        build-essential cmake git sed \
    && git clone --depth 1 https://github.com/LibreDWG/libredwg.git /tmp/libredwg \
    && cd /tmp/libredwg && mkdir build && cd build \
    && cmake .. -DCMAKE_INSTALL_PREFIX=/usr/local \
        -DCMAKE_C_FLAGS="-w" \
    && sed -i 's/-Werror//g' CMakeCache.txt CMakeFiles/*.dir/flags.make 2>/dev/null || true \
    && make -j$(nproc) CFLAGS="-w" \
    && make install \
    && ldconfig \
    && rm -rf /tmp/libredwg \
    && apt-get purge -y build-essential cmake git \
    && apt-get autoremove -y \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY . .

RUN mkdir -p uploads outputs

EXPOSE 10000

CMD ["gunicorn", "app:app", "--bind", "0.0.0.0:10000", "--timeout", "300", "--workers", "2"]
