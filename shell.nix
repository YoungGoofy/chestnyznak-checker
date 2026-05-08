# Стандартный shell.nix (для nix-shell без flakes)
let
  pkgs = import <nixpkgs> { };
  pythonWithTk = pkgs.python311.withPackages (ps: with ps; [
    tkinter
    openpyxl
    python-dotenv
  ]);
in pkgs.mkShell {
  buildInputs = [ pythonWithTk ];
  shellHook = ''
    echo "🐍 Python $(python3 --version) + tkinter + openpyxl"
    echo "🚀 Запуск GUI: python gui_app.py"
  '';
}
