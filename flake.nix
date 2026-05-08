{
  description = "Честный Знак — Проверка кодов маркировки";
  inputs = { nixpkgs.url = "github:NixOS/nixpkgs/nixos-unstable"; };

  outputs = { self, nixpkgs }:
    let
      supported = [ "x86_64-linux" "x86_64-darwin" "aarch64-darwin" "aarch64-linux" ];
      forAllSystems = nixpkgs.lib.genAttrs supported;
    in {
      devShells = forAllSystems (system:
        let
          pkgs = import nixpkgs { inherit system; };
          pythonWithTk = pkgs.python311.withPackages (ps: with ps; [
            tkinter
            openpyxl
            python-dotenv
          ]);
        in {
          default = pkgs.mkShell {
            buildInputs = [
              pythonWithTk
            ];
            shellHook = ''
              echo "🐍 Python $(python3 --version) + tkinter + openpyxl"
              echo "🚀 Запуск GUI: python gui_app.py"
              echo "📋 Запуск CLI: python check_codes.py --true -f codes.txt -o result.xlsx"
            '';
          };
        }
      );
    };
}
