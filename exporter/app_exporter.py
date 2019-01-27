import argparse
import logging
import os, pathlib, platform, shutil, sys, time
import keyring, sharepy
from datetime import datetime


FULL_DATETIME_FORMAT = "%Y-%m-%d %H:%M:%S"
YEAR_MONTH_FORMAT = "%Y-%m"


def main():
    start_timestamp = time.time()
    logging.info("Starting up (platform: %s, current timestamp: %s (UTC: %s)",
                 platform.system(), start_timestamp, timestamp_to_string(start_timestamp, FULL_DATETIME_FORMAT))
    args = parse_args()
    logging.info("Arguments: %s", args)

    # list files
    filenames = find_files(args.input_dir)
    logging.info("Found %s files under input directory: %s", len(filenames), args.input_dir)

    # choose files meeting criteria (i.e. old files)
    maximum_timestamp = start_timestamp - days_to_seconds(args.days)
    logging.info("Filtering out files with modification date after %f (UTC: %s)",
                 maximum_timestamp, timestamp_to_string(maximum_timestamp, FULL_DATETIME_FORMAT))
    oldfile_names = filter_old_files(filenames, maximum_timestamp)
    logging.info("Found %s files modified more than %s days ago", len(oldfile_names), args.days)

    if args.dry_run:
        logging.info("DRY RUN mode - listing files")
        logging.info([str(filename_object) for filename_object in oldfile_names])
        logging.info("DRY RUN mode - quitting now")
        sys.exit(0)

    # move files to working folder
    os.makedirs(args.working_dir, exist_ok=True)
    move_files(oldfile_names, args.input_dir, args.working_dir)

    # upload from working folder to Sharepoint
    selected_filenames = find_files(args.working_dir)
    logging.info("Found %s files under working directory: %s", len(selected_filenames), args.working_dir)
    uploaded_filenames = upload_to_sharepoint(
        selected_filenames, args.sharepoint_host, args.sharepoint_site, args.sharepoint_library, args.user
    )

    # delete uploaded files from working folder
    if args.delete:
        logging.info("%s files uploaded successfully will be deleted now", len(uploaded_filenames))
        delete_files(uploaded_filenames)
    else:
        logging.info("NO DELETE mode - uploaded files will remain in working folder")

    logging.info("All done - thank you and bye bye!")


def find_files(dir_name):
    filenames = []

    for path, subdirs, files in os.walk(dir_name):
        for name in files:
            filenames.append(pathlib.PurePath(path, name))

    return filenames


def filter_old_files(filenames, maximum_timestamp):
    return [filename for filename in filenames if modification_date(filename) < maximum_timestamp]


def move_files(filenames, base_dir, target_dir):
    for filename in filenames:
        filename_object = pathlib.PurePath(filename)

        base_filename = filename_object.name
        file_rel_directory = filename_object.parent.relative_to(base_dir)

        file_target_directory = pathlib.Path(target_dir, file_rel_directory)
        if not file_target_directory.exists():
            file_target_directory.mkdir(parents=True)
        target_filename = pathlib.PurePath(file_target_directory, base_filename)
        shutil.move(filename, target_filename)


def upload_to_sharepoint(filenames, sharepoint_host, sharepoint_site, sharepoint_library, user):
    logging.info("Using SharePoint host %s, site %s, library %s", sharepoint_host, sharepoint_site, sharepoint_library)
    password = keyring.get_password(sharepoint_host, user)
    if password is None:
        raise ValueError(f"Could not get password from keyring for {user}@{sharepoint_host}")

    s = sharepy.connect(f"https://{sharepoint_host}", username=user, password=password)

    if s.get(f"https://{sharepoint_host}").status_code == 403:
        raise ValueError(f"Forbidden for {sharepoint_host} - authentication failed")

    headers = {"accept": "application/json;odata=verbose",
           "content-type": "application/x-www-urlencoded; charset=UTF-8"}

    successful_upload_count = 0
    failed_upload_count = 0
    created_folders = set()
    uploaded_filenames = []
    for filename in filenames:
        filename_object = pathlib.Path(filename)
        with open(filename_object, 'rb') as read_file:
            content = read_file.read()
        folder_name = timestamp_to_string(modification_date(filename), YEAR_MONTH_FORMAT)
        relative_target = sharepoint_library +'/' + folder_name

        if relative_target not in created_folders:
            p = s.post(f"https://{sharepoint_host}/sites/{sharepoint_site}/_api/web/folders",
                json={
                    "__metadata": { "type": "SP.Folder" },
                    "ServerRelativeUrl": relative_target
                })
            if p.status_code < 200 or p.status_code >= 300:
                logging.error("ERROR: Post to create folder %s resulted in status %s",
                              relative_target, p.status_code)
                failed_upload_count += 1
                continue
            else:
                created_folders.add(relative_target)
                logging.info("OK: Post to create folder %s resulted in status %s", relative_target, p.status_code)

        p = s.post(f"https://{sharepoint_host}/sites/{sharepoint_site}/_api/web/getFolderByServerRelativeUrl('{sharepoint_library}/{folder_name}')/Files//add(url='{filename_object.name}', overwrite=true)", data=content, headers=headers)
        if p.status_code < 200 or p.status_code >= 300:
            logging.error("ERROR: Post for file %s resulted in status %s", filename, p.status_code)
            failed_upload_count += 1
        else:
            logging.info("OK: Post for file %s resulted in status %s", filename, p.status_code)
            successful_upload_count += 1
            uploaded_filenames.append(filename)

    logging.info("Successful uploads: %s, failed uploads: %s", successful_upload_count, failed_upload_count)
    return uploaded_filenames


def delete_files(filenames):
    for filename in filenames:
        pathlib.Path(filename).unlink()


def timestamp_to_string(timestamp, date_format):
    return datetime.utcfromtimestamp(timestamp).strftime(date_format)


def days_to_seconds(days):
    return float(days) * 60 * 60 * 24


def modification_date(path_to_file):
    return os.path.getmtime(path_to_file)


def parse_args():
    parser = argparse.ArgumentParser(description='Upload files that meet criteria to SharePoint.')
    parser.add_argument('--input_dir', '-i', required=True, type=str, help='Directory to look for files in')
    parser.add_argument('--working_dir', '-w', required=True, type=str, help='Directory to use as working directory')
    parser.add_argument('--sharepoint_host', '-sh', required=True, type=str,
                        help='SharePoint host (e.g. example.sharepoint.com)')
    parser.add_argument('--sharepoint_site', '-ss', required=True, type=str,
                        help='Name of SharePoint site to upload files to')
    parser.add_argument('--sharepoint_library', '-sl', required=True, type=str,
                        help='Name of SharePoint library to upload files to')
    parser.add_argument('--days', '-d', required=True, type=str,
                        help='Only files modified more than that many days are uploaded')
    parser.add_argument('--user', '-u', required=True, type=str, help='SharePoint user ID')

    delete_parser = parser.add_mutually_exclusive_group(required=False)
    delete_parser.add_argument('--delete', dest='delete', action='store_true')
    delete_parser.add_argument('--no-delete', dest='delete', action='store_false')
    parser.set_defaults(delete=True)

    dryrun_parser = parser.add_mutually_exclusive_group(required=False)
    dryrun_parser.add_argument('--dry-run', dest='dry_run', action='store_true')
    dryrun_parser.add_argument('--no-dry-run', dest='dry_run', action='store_false')
    parser.set_defaults(dry_run=False)

    args = parser.parse_args()
    return args


if __name__ == '__main__':
    logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(name)s - %(levelname)s - %(message)s')
    main()
