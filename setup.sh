#!/usr/bin/env bash
set -euo pipefail

# Install .NET 9 SDK if not present
DOTNET_ROOT="$HOME/.dotnet"
export DOTNET_ROOT
export PATH="$DOTNET_ROOT:$PATH"

if ! command -v dotnet >/dev/null || ! dotnet --version | grep -q '^9'; then
  curl -sSL https://dot.net/v1/dotnet-install.sh | bash /dev/stdin --channel 9.0 --install-dir "$DOTNET_ROOT"
fi

dotnet --info

# Move to repository root
cd "$(dirname "$0")" || exit 1


echo 'export DOTNET_ROOT="$HOME/.dotnet"' >> ~/.bashrc
echo 'export PATH="$DOTNET_ROOT:$PATH"' >> ~/.bashrc

