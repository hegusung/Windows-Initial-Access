import os
import yaml

def enumerate_lolbas(lolbas_path):
    # Iterate over files in directory
    for path, folders, files in os.walk(lolbas_path):
        # Open file
        for filename in files:
            if filename.endswith(".yml"):
                with open(os.path.join(lolbas_path, filename)) as f:
                    data = f.read()
                    yield data

        # List contain of folder
        for folder_name in folders:
            folder_path = os.path.join(lolbas_path, folder_name)
            for yml_data in enumerate_lolbas(folder_path):
                yield yml_data

        break


def list_tags(lolbas_path):
    tags = []

    for yml_data in enumerate_lolbas(lolbas_path):
        try:
            yml_dict = yaml.safe_load(yml_data)
            for command in yml_dict['Commands']:
                if 'Tags' in command:
                    for t in command['Tags']:
                        tags.append("%s: %s" % (list(t.keys())[0], list(t.values())[0]))
        except yaml.YAMLError as exc:
            print(exc)

    tags = list(set(tags))

    return tags

def lolbas_to_json(lolbas_path):
    json = {}

    for yml_data in enumerate_lolbas(lolbas_path):
        try:
            yml_dict = yaml.safe_load(yml_data)

            name = yml_dict['Name']

            paths = []
            for p_dict in yml_dict['Full_Path']:
                p = p_dict['Path']

                if p.lower().startswith('c:\\'):
                    paths.append(p)
                else:
                    print("[%s] Invalid path: %s" % (name, p))

            if len(paths) == 0:
                print("[%s] No paths to check" % (name, ))
                continue


            tags = []

            for command in yml_dict['Commands']:
                if 'Tags' in command:
                    for t in command['Tags']:
                        tags.append("%s: %s" % (list(t.keys())[0], list(t.values())[0]))
                else:
                    tags.append(command['Category'])

            tags = list(set(tags))

            lolbas_data = {
                    "paths": paths,
                    "tags": tags,
            }
            json[name] = lolbas_data
        except yaml.YAMLError as exc:
            print(exc)

    return json

