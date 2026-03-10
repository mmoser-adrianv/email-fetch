"""add_teams_attachments

Revision ID: e5f6a7b8c9d0
Revises: d4e5f6a7b8c9
Create Date: 2026-03-09 00:00:00.000000

"""
from alembic import op
import sqlalchemy as sa


# revision identifiers, used by Alembic.
revision = 'e5f6a7b8c9d0'
down_revision = 'd4e5f6a7b8c9'
branch_labels = None
depends_on = None


def upgrade():
    op.add_column('teams_chats', sa.Column('media_status', sa.String(length=32), nullable=True))

    op.create_table(
        'teams_attachments',
        sa.Column('id', sa.Integer(), nullable=False),
        sa.Column('teams_message_id', sa.Integer(), nullable=False),
        sa.Column('attachment_id', sa.String(length=512), nullable=True),
        sa.Column('attachment_type', sa.String(length=32), nullable=True),
        sa.Column('name', sa.String(length=512), nullable=True),
        sa.Column('content_type', sa.String(length=255), nullable=True),
        sa.Column('content_url', sa.Text(), nullable=True),
        sa.Column('local_path', sa.String(length=1024), nullable=True),
        sa.Column('download_status', sa.String(length=32), nullable=True),
        sa.Column('created_at', sa.DateTime(), server_default=sa.text('now()'), nullable=True),
        sa.ForeignKeyConstraint(['teams_message_id'], ['teams_messages.id'], ),
        sa.PrimaryKeyConstraint('id'),
    )
    op.create_index('ix_teams_attachments_teams_message_id', 'teams_attachments', ['teams_message_id'])


def downgrade():
    op.drop_index('ix_teams_attachments_teams_message_id', table_name='teams_attachments')
    op.drop_table('teams_attachments')
    op.drop_column('teams_chats', 'media_status')
