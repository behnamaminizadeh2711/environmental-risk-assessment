import brightway2 as bw

# حذف پروژه فعلی
project_name = "EnvironmentalRiskProject"
if project_name in bw.projects:
    bw.projects.delete_project(project_name, delete_dir=True)
    print(f"پروژه {project_name} با موفقیت حذف شد.")
else:
    print(f"پروژه {project_name} وجود ندارد.")