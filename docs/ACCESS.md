# Access model

The GitHub repository and GCP Artifact Registry repository are private.

## Buyer access

Grant a buyer read-only access as an outside collaborator or through a
buyer-specific GitHub team with repository `Read` permission. A read-only
outside collaborator cannot manage repository access or inspect the private
repository's full collaborator list. Do not grant `Admin` or organization
owner privileges.

For the container image, grant `roles/artifactregistry.reader` only to the
buyer's named Google principal and only on the `alder-ridge-samples` Artifact
Registry repository. Repository collaborators and GCP image readers are
managed separately.

## Revocation

Remove the GitHub collaborator/team grant and the GCP repository-level IAM
binding. No public repository, public bucket, or anonymous image-pull grant is
used by this package.

## Secrets

The repository and image contain no HUD, OpenAI, Anthropic, GitHub, or GCP
credentials. Runtime semantic-review credentials, if used, must be injected
by the operator and are removed before the agent-facing shell starts.
