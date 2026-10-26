import * as React from "react";
import {
	DefaultButton,
	Dialog,
	DialogFooter,
	DialogType,
	Dropdown,
	IDropdownOption,
	Label,
	MessageBar,
	MessageBarType,
	PrimaryButton,
	SearchBox,
	Spinner,
	SpinnerSize,
} from "@fluentui/react";
import * as strings from "CrscPromptsWebPartStrings";
import { ICrscPrompt } from "../models/ICrscPrompt";
import {
	filterAndSortPrompts,
	formatModifiedDate,
	IPromptSortOrder,
} from "../utils/promptUtils";
import { ICrscPromptsProps } from "./ICrscPromptsProps";
import styles from "./CrscPrompts.module.scss";

// Returns page numbers to display; 0 marks an ellipsis gap.
function getPageItems(current: number, count: number): number[] {
	const pages: number[] = [];
	for (let p = 1; p <= count; p++) {
		if (p === 1 || p === count || Math.abs(p - current) <= 1) {
			pages.push(p);
		} else if (pages[pages.length - 1] !== 0) {
			pages.push(0);
		}
	}
	return pages;
}

interface ICrscPromptsState {
	prompts: ICrscPrompt[];
	loading: boolean;
	error?: string;
	query: string;
	deliverable: string;
	category: string;
	sortOrder: IPromptSortOrder;
	page: number;
	selectedPrompt?: ICrscPrompt;
}

export default class CrscPrompts extends React.Component<
	ICrscPromptsProps,
	ICrscPromptsState
> {
	public constructor(props: ICrscPromptsProps) {
		super(props);
		this.state = {
			prompts: [],
			loading: !!props.listId,
			query: "",
			deliverable: "",
			category: "",
			sortOrder: "newest",
			page: 1,
		};
	}

	public componentDidMount(): void {
		if (this.props.listId) {
			this.loadPrompts();
		}
	}

	public componentDidUpdate(previousProps: ICrscPromptsProps): void {
		if (previousProps.listId !== this.props.listId) {
			this.setState(
				{
					prompts: [],
					loading: !!this.props.listId,
					error: undefined,
					page: 1,
					selectedPrompt: undefined,
				},
				() => {
					if (this.props.listId) {
						this.loadPrompts();
					}
				},
			);
		}
	}

	private loadPrompts = (): void => {
		if (!this.props.listId) {
			return;
		}

		this.setState({ loading: true, error: undefined });
		this.props.service
			.getPrompts(this.props.listId)
			.then((prompts) => this.setState({ prompts, loading: false }))
			.catch((error: Error) =>
				this.setState({
					loading: false,
					error: error.message || strings.LoadError,
				}),
			);
	};

	private onSearch = (
		_event?: React.ChangeEvent<HTMLInputElement>,
		value?: string,
	): void => {
		this.setState({ query: value || "", page: 1 });
	};

	private onDeliverableChange = (
		_event: React.FormEvent<HTMLDivElement>,
		option?: IDropdownOption,
	): void => {
		this.setState({ deliverable: option ? String(option.key) : "", page: 1 });
	};

	private onCategoryChange = (
		_event: React.FormEvent<HTMLDivElement>,
		option?: IDropdownOption,
	): void => {
		this.setState({ category: option ? String(option.key) : "", page: 1 });
	};

	private onSortChange = (
		_event: React.FormEvent<HTMLDivElement>,
		option?: IDropdownOption,
	): void => {
		if (option) {
			this.setState({ sortOrder: option.key as IPromptSortOrder, page: 1 });
		}
	};

	private clearFilters = (): void => {
		this.setState({
			query: "",
			deliverable: "",
			category: "",
			sortOrder: "newest",
			page: 1,
		});
	};

	private renderPromptDetails = (prompt: ICrscPrompt): React.ReactElement => {
		const fields: { label: string; value: string }[] = [
			{ label: strings.PromptTitle, value: prompt.Title },
			{ label: strings.PromptInput, value: prompt.Input },
			{ label: strings.AITool, value: prompt.AITool },
			{ label: strings.Category, value: prompt.Category },
			{ label: strings.Deliverable, value: prompt.Deliverable },
			{ label: strings.RiskLevel, value: prompt.RiskLevel },
			{ label: strings.LastModified, value: formatModifiedDate(prompt.Modified) },
			{ label: strings.SourceInputs, value: prompt.SourceInputsRequired },
			{ label: strings.Limitations, value: prompt.LimitationsandConstraints },
			{ label: strings.AdditionalInformation, value: prompt.AdditionalInformation },
		];

		return (
			<div className={styles.promptDetails}>
				{fields.map((field) => (
					<div className={styles.detailField} key={field.label}>
						<Label>{field.label}</Label>
						<div className={styles.detailValue}>{field.value || strings.NotProvided}</div>
					</div>
				))}
			</div>
		);
	};

	public render(): React.ReactElement<ICrscPromptsProps> {
		const {
			listId,
			itemsPerPage,
			modalMinWidth,
			clearAllButtonColor,
			cardBorderColor,
			cardTitleColor,
			viewButtonTextColor,
			viewButtonBackgroundColor,
			onConfigure,
		} = this.props;
		const {
			prompts,
			loading,
			error,
			query,
			deliverable,
			category,
			sortOrder,
			page,
			selectedPrompt,
		} = this.state;

		if (!listId) {
			return (
				<section className={styles.crscPrompts}>
					<MessageBar messageBarType={MessageBarType.info}>
						{strings.ConfigureMessage}
					</MessageBar>
					<PrimaryButton className={styles.configureButton} onClick={onConfigure}>
						{strings.ConfigureButton}
					</PrimaryButton>
				</section>
			);
		}

		const filteredPrompts = filterAndSortPrompts(
			prompts,
			query,
			deliverable,
			category,
			sortOrder,
		);
		const pageCount = Math.max(1, Math.ceil(filteredPrompts.length / itemsPerPage));
		const currentPage = Math.min(page, pageCount);
		const startIndex = (currentPage - 1) * itemsPerPage;
		const visiblePrompts = filteredPrompts.slice(
			startIndex,
			startIndex + itemsPerPage,
		);
		const deliverables = prompts
			.map((prompt) => prompt.Deliverable)
			.filter((value, index, values) => !!value && values.indexOf(value) === index)
			.sort();
		const categories = prompts
			.map((prompt) => prompt.Category)
			.filter((value, index, values) => !!value && values.indexOf(value) === index)
			.sort();
		const sortOptions: IDropdownOption[] = [
			{ key: "newest", text: strings.Newest },
			{ key: "oldest", text: strings.Oldest },
			{ key: "aToZ", text: strings.AToZ },
			{ key: "zToA", text: strings.ZToA },
		];

		const pagination = (
			<nav className={styles.pagination} aria-label={strings.Pagination}>
				<DefaultButton
					className={styles.pageButton}
					ariaLabel={strings.Previous}
					disabled={currentPage <= 1}
					onClick={() => this.setState({ page: currentPage - 1 })}>
					&lt;
				</DefaultButton>
				{getPageItems(currentPage, pageCount).map((item, index) =>
					item === 0 ? (
						<span key={`gap-${index}`}>…</span>
					) : (
						<DefaultButton
							key={item}
							className={styles.pageButton}
							checked={item === currentPage}
							ariaLabel={strings.PageOf.replace("{0}", String(item)).replace(
								"{1}",
								String(pageCount),
							)}
							onClick={() => this.setState({ page: item })}>
							{item}
						</DefaultButton>
					),
				)}
				<DefaultButton
					className={styles.pageButton}
					ariaLabel={strings.Next}
					disabled={currentPage >= pageCount}
					onClick={() => this.setState({ page: currentPage + 1 })}>
					&gt;
				</DefaultButton>
			</nav>
		);

		return (
			<section className={styles.crscPrompts}>
				<div className={styles.filterPanel}>
					<div className={styles.searchRow}>
						<SearchBox
							className={styles.searchBox}
							placeholder={strings.SearchPlaceholder}
							ariaLabel={strings.SearchPlaceholder}
							value={query}
							onChange={this.onSearch}
						/>
						<DefaultButton
							onClick={this.clearFilters}
							styles={{
								root: {
									backgroundColor: clearAllButtonColor,
									borderColor: clearAllButtonColor,
								},
								rootHovered: {
									backgroundColor: clearAllButtonColor,
									borderColor: clearAllButtonColor,
								},
								rootPressed: {
									backgroundColor: clearAllButtonColor,
									borderColor: clearAllButtonColor,
								},
							}}>
							{strings.ClearAll}
						</DefaultButton>
					</div>
					<div className={styles.filterRow}>
						<Label className={styles.filtersLabel}>{strings.Filters}</Label>
						<Dropdown
							className={styles.filterDropdown}
							label={""}
							selectedKey={deliverable || undefined}
							options={[
								{ key: "", text: strings.AllDeliverables },
								...deliverables.map((value) => ({ key: value, text: value })),
							]}
							onChange={this.onDeliverableChange}
						/>
						<Dropdown
							className={styles.filterDropdown}
							label={""}
							selectedKey={category || undefined}
							options={[
								{ key: "", text: strings.AllCategories },
								...categories.map((value) => ({ key: value, text: value })),
							]}
							onChange={this.onCategoryChange}
						/>
						<div className={styles.sortGroup}>
							<Label className={styles.filtersLabel} htmlFor="crscPromptsSort">
								{strings.SortBy}
							</Label>
							<Dropdown
								id="crscPromptsSort"
								className={styles.sortDropdown}
								selectedKey={sortOrder}
								options={sortOptions}
								onChange={this.onSortChange}
							/>
						</div>
					</div>
				</div>

				{loading ? (
					<Spinner
						className={styles.statusMessage}
						size={SpinnerSize.medium}
						label={strings.Loading}
					/>
				) : error ? (
					<MessageBar
						className={styles.statusMessage}
						messageBarType={MessageBarType.error}
						actions={
							<PrimaryButton onClick={this.loadPrompts}>{strings.Retry}</PrimaryButton>
						}>
						{error}
					</MessageBar>
				) : filteredPrompts.length === 0 ? (
					<MessageBar
						className={styles.statusMessage}
						messageBarType={MessageBarType.info}>
						{strings.NoPrompts}
					</MessageBar>
				) : (
					<>
						<div className={styles.resultsBar}>
							<span>
								{strings.ShowingEntries.replace("{0}", String(startIndex + 1))
									.replace(
										"{1}",
										String(Math.min(startIndex + itemsPerPage, filteredPrompts.length)),
									)
									.replace("{2}", String(filteredPrompts.length))}
							</span>
							{pagination}
						</div>
						<div className={styles.promptGrid}>
							{visiblePrompts.map((prompt) => (
								<article
									className={styles.promptCard}
									key={prompt.Id}
									style={{ borderColor: cardBorderColor }}>
									<div
										className={styles.deliverableTag}
										style={{ backgroundColor: cardBorderColor }}>
										{prompt.Deliverable || strings.NotProvided}
									</div>
									<div className={styles.promptContent}>
										<h2 className={styles.promptTitle} style={{ color: cardTitleColor }}>
											{prompt.Title}
										</h2>
										<div className={styles.paddingBottomMedium}>
											<strong>{strings.Category}:</strong> {prompt.Category || strings.NotProvided}
										</div>
										<div>
											<strong>{strings.AITool}:</strong> {prompt.AITool || strings.NotProvided}
										</div>
										<div className={styles.paddingBottomMedium}>
											<strong>{strings.RiskLevel}:</strong> {prompt.RiskLevel || strings.NotProvided}
										</div>
										<div className={styles.paddingBottomMedium}>
											<strong>{strings.LastModified}:</strong> {formatModifiedDate(prompt.Modified)}
										</div>
										<PrimaryButton
											className={styles.viewButton}
											styles={{
												root: {
													backgroundColor: viewButtonBackgroundColor,
													borderColor: viewButtonBackgroundColor,
													color: viewButtonTextColor,
												},
												rootHovered: {
													backgroundColor: viewButtonBackgroundColor,
													borderColor: viewButtonBackgroundColor,
													color: viewButtonTextColor,
												},
											}}
											onClick={() => this.setState({ selectedPrompt: prompt })}>
											{strings.ViewFullPrompt}
										</PrimaryButton>
									</div>
								</article>
							))}
						</div>
						<div className={styles.bottomPagination}>{pagination}</div>
					</>
				)}

				<Dialog
					hidden={!selectedPrompt}
					onDismiss={() => this.setState({ selectedPrompt: undefined })}
					dialogContentProps={{
						type: DialogType.largeHeader,
						title: selectedPrompt ? selectedPrompt.Title : "",
						closeButtonAriaLabel: strings.Close,
					}}
					modalProps={{ isBlocking: false }}
					styles={{
						main: {
							width: "calc(100vw - 32px) !important",
							minWidth: `${modalMinWidth}px !important`,
							maxWidth: "none !important",
						},
					}}>
					{selectedPrompt && this.renderPromptDetails(selectedPrompt)}
					<DialogFooter>
						<DefaultButton onClick={() => this.setState({ selectedPrompt: undefined })}>
							{strings.Close}
						</DefaultButton>
					</DialogFooter>
				</Dialog>
			</section>
		);
	}
}
