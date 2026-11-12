CREATE TABLE "pending_ops" (
	"id" serial PRIMARY KEY,
	"kind" text NOT NULL,
	"action" text NOT NULL,
	"row" integer,
	"expect" text,
	"base" jsonb,
	"rec" jsonb DEFAULT '{}' NOT NULL,
	"created_by" text,
	"created_at" timestamp DEFAULT now()
);
