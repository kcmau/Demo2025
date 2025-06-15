from flask_sqlalchemy import SQLAlchemy
from datetime import datetime

db = SQLAlchemy()

class FormSubmission(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    name = db.Column(db.String(100), nullable=False)
    school_name = db.Column(db.String(200), nullable=False)
    selected_robots = db.Column(db.Text, nullable=False)  # Store as comma-separated string
    submitted_at = db.Column(db.DateTime, default=datetime.utcnow)

    def __repr__(self):
        return f'<FormSubmission {self.name} - {self.school_name}>'

    def to_dict(self):
        return {
            'id': self.id,
            'name': self.name,
            'school_name': self.school_name,
            'selected_robots': self.selected_robots.split(',') if self.selected_robots else [],
            'submitted_at': self.submitted_at.isoformat()
        }

class SubmissionCounter(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    count = db.Column(db.Integer, default=0)
    max_submissions = db.Column(db.Integer, default=10)

    def __repr__(self):
        return f'<SubmissionCounter {self.count}/{self.max_submissions}>'

    def to_dict(self):
        return {
            'id': self.id,
            'count': self.count,
            'max_submissions': self.max_submissions,
            'submissions_remaining': self.max_submissions - self.count,
            'is_limit_reached': self.count >= self.max_submissions
        }

