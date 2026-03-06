"""add body_text to emails

Revision ID: a1b2c3d4e5f6
Revises: 07bbe8764531
Create Date: 2026-03-06 12:00:00.000000

"""
from alembic import op
import sqlalchemy as sa


# revision identifiers, used by Alembic.
revision = 'a1b2c3d4e5f6'
down_revision = '07bbe8764531'
branch_labels = None
depends_on = None


def upgrade():
    op.add_column('emails', sa.Column('body_text', sa.Text(), nullable=True))


def downgrade():
    op.drop_column('emails', 'body_text')
