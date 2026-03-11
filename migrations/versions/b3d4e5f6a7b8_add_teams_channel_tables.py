"""add_teams_channel_tables

Revision ID: b3d4e5f6a7b8
Revises: a1b2c3d4e5f7
Create Date: 2026-03-10 00:00:00.000000

"""
from alembic import op
import sqlalchemy as sa


# revision identifiers, used by Alembic.
revision = 'b3d4e5f6a7b8'
down_revision = 'a1b2c3d4e5f7'
branch_labels = None
depends_on = None


def upgrade():
    op.create_table(
        'teams_teams',
        sa.Column('id', sa.Integer(), nullable=False),
        sa.Column('team_id', sa.String(length=512), nullable=False),
        sa.Column('display_name', sa.String(length=512), nullable=True),
        sa.Column('description', sa.Text(), nullable=True),
        sa.Column('scraped_at', sa.DateTime(), nullable=True),
        sa.Column('created_at', sa.DateTime(), server_default=sa.text('now()'), nullable=True),
        sa.PrimaryKeyConstraint('id'),
        sa.UniqueConstraint('team_id'),
    )
    op.create_index('ix_teams_teams_team_id', 'teams_teams', ['team_id'])

    op.create_table(
        'teams_channels',
        sa.Column('id', sa.Integer(), nullable=False),
        sa.Column('teams_team_id', sa.Integer(), nullable=False),
        sa.Column('channel_id', sa.String(length=512), nullable=False),
        sa.Column('display_name', sa.String(length=512), nullable=True),
        sa.Column('description', sa.Text(), nullable=True),
        sa.Column('scraped_at', sa.DateTime(), nullable=True),
        sa.Column('created_at', sa.DateTime(), server_default=sa.text('now()'), nullable=True),
        sa.ForeignKeyConstraint(['teams_team_id'], ['teams_teams.id']),
        sa.PrimaryKeyConstraint('id'),
        sa.UniqueConstraint('channel_id'),
    )
    op.create_index('ix_teams_channels_channel_id', 'teams_channels', ['channel_id'])

    op.create_table(
        'teams_channel_posts',
        sa.Column('id', sa.Integer(), nullable=False),
        sa.Column('teams_channel_id', sa.Integer(), nullable=False),
        sa.Column('message_id', sa.String(length=512), nullable=False),
        sa.Column('sender_name', sa.String(length=255), nullable=True),
        sa.Column('sender_email', sa.String(length=255), nullable=True),
        sa.Column('subject', sa.String(length=512), nullable=True),
        sa.Column('content_html', sa.Text(), nullable=True),
        sa.Column('content_text', sa.Text(), nullable=True),
        sa.Column('created_date_time', sa.DateTime(timezone=True), nullable=True),
        sa.Column('importance', sa.String(length=32), nullable=True),
        sa.Column('web_url', sa.Text(), nullable=True),
        sa.Column('message_type', sa.String(length=64), nullable=True),
        sa.Column('created_at', sa.DateTime(), server_default=sa.text('now()'), nullable=True),
        sa.ForeignKeyConstraint(['teams_channel_id'], ['teams_channels.id']),
        sa.PrimaryKeyConstraint('id'),
        sa.UniqueConstraint('message_id'),
    )
    op.create_index('ix_teams_channel_posts_message_id', 'teams_channel_posts', ['message_id'])


def downgrade():
    op.drop_index('ix_teams_channel_posts_message_id', table_name='teams_channel_posts')
    op.drop_table('teams_channel_posts')
    op.drop_index('ix_teams_channels_channel_id', table_name='teams_channels')
    op.drop_table('teams_channels')
    op.drop_index('ix_teams_teams_team_id', table_name='teams_teams')
    op.drop_table('teams_teams')
