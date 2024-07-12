import subprocess


def update_requirements():
    # Run pip list --outdated to get a list of outdated packages
    result = subprocess.run(['pip', 'list', '--outdated'], stdout=subprocess.PIPE, text=True)
    outdated_packages = result.stdout.splitlines()

    updated_lines = []
    with open('requirements.txt', 'r') as file:
        lines = file.readlines()
        for line in lines:
            package, current_version = line.strip().split('==')
            for outdated in outdated_packages:
                if package in outdated:
                    _, _, latest_version = outdated.split()
                    updated_lines.append(f"{package}=={latest_version}\n")
                    break
            else:
                updated_lines.append(line)

    with open('requirements.txt', 'w') as file:
        file.writelines(updated_lines)


if __name__ == "__main__":
    update_requirements()
