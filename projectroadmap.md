https://github.com/yas0h00/job-application-tracker 

After reviewing the base code above, I found that it uses

- Flask for the web framework
- SQLAlchemy for database models
- SQLite support
- Flask-Login for authentication
- Flask-Migrate for database migrations

Some of the strengths include that the project is closely related to my own idea, and it solves a very similar problem, and it shows different things such as

- how to structure a Flask project
- how to design an applications table/model
- how to display dashboard statistics
- how to organize templates and static files

The main weakness is that it is more complex than what I believe I could do and what I need for my project. It has things such as multi-user login and signup page. It has user set and email set up, as well as PDF and CVS export.

Progress so far. 
- [x] Added Flask startup block 
- [x] Fixed app issue
- [x] Created the app with create_app()
- [x] Connected the dashboard route to the main index.html template
- [x] Fixed Jinja error, was missing {% endblock %}
- [x] Tested the app can run
      
Currently working on
- [ ] Using the Excel import with my internship applications data
- [ ] Making sure the dashboard updates after import
- [ ] Cleaning up overall UI 
