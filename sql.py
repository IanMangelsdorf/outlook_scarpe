from sqlalchemy import create_engine
from sqlalchemy import MetaData
from sqlalchemy.ext.declarative import declarative_base
from sqlalchemy import Column, Integer, String, Boolean, Float, Date
from sqlalchemy import ForeignKey
from sqlalchemy.orm import relationship





Base = declarative_base()

metadata = MetaData()





class Emails(Base):
    __tablename__ = 'email'
    id = Column(Integer, primary_key=True, nullable=False)
    subject= Column(String(255), nullable=True, unique=False)
    sender = Column(String(255), nullable=False, unique=False)
    sender_email = Column(String(255), nullable = True,unique = False)
    date = Column(String(50),nullable = True,unique = False)
    body = Column(String)
    folder = Column(String)
    subfolder = Column(String)
    company = Column(String(50), nullable = True,unique = False)

    def __init__(self,subject,sender, email, date, body, fold, subfold,company):
        self.subject = subject
        self.sender = sender
        self.sender_email = email
        self.date = date
        self.body= body
        self.folder=fold
        self.subject= subfold
        self.company= company


    def __repr__(self):
        return '<Sample %r, %r' % (
            self.subject, self.sender
        )



engine = create_engine('sqlite:///outlook.db')
Base.metadata.create_all(engine)

