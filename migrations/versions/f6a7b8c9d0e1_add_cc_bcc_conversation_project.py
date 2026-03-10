"""add_cc_bcc_conversation_project

Revision ID: f6a7b8c9d0e1
Revises: e5f6a7b8c9d0
Create Date: 2026-03-09 00:00:00.000000

"""
from alembic import op
import sqlalchemy as sa


# revision identifiers, used by Alembic.
revision = 'f6a7b8c9d0e1'
down_revision = 'e5f6a7b8c9d0'
branch_labels = None
depends_on = None


def upgrade():
    op.create_table(
        'projects',
        sa.Column('id', sa.Integer(), nullable=False),
        sa.Column('title', sa.String(length=255), nullable=False),
        sa.Column('project_number', sa.String(length=64), nullable=True),
        sa.Column('created_at', sa.DateTime(), server_default=sa.text('now()'), nullable=True),
        sa.PrimaryKeyConstraint('id'),
    )

    op.add_column('emails', sa.Column('to_recipients', sa.Text(), nullable=True))
    op.add_column('emails', sa.Column('cc_recipients', sa.Text(), nullable=True))
    op.add_column('emails', sa.Column('bcc_recipients', sa.Text(), nullable=True))
    op.add_column('emails', sa.Column('conversation_id', sa.String(length=512), nullable=True))
    op.create_index('ix_emails_conversation_id', 'emails', ['conversation_id'])
    op.add_column('emails', sa.Column('project_id', sa.Integer(), nullable=True))
    op.create_foreign_key('fk_emails_project_id', 'emails', 'projects', ['project_id'], ['id'])


def downgrade():
    op.drop_constraint('fk_emails_project_id', 'emails', type_='foreignkey')
    op.drop_column('emails', 'project_id')
    op.drop_index('ix_emails_conversation_id', table_name='emails')
    op.drop_column('emails', 'conversation_id')
    op.drop_column('emails', 'bcc_recipients')
    op.drop_column('emails', 'cc_recipients')
    op.drop_column('emails', 'to_recipients')
    op.drop_table('projects')
